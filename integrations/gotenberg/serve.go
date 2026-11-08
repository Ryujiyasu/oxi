// SPDX-License-Identifier: MIT OR Apache-2.0

package oxi

import (
	"bufio"
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"log/slog"
	"os/exec"
	"sync"
	"time"
)

// errServerGone is returned for requests in flight when the oxidocs process
// exits, whatever the reason (crash, abort, kill after a timeout). The
// request itself may be fine: another one on the same process may have
// brought it down.
var errServerGone = errors.New("oxidocs serve exited")

// errTimeout is returned for the request that ran past its time limit; its
// process is replaced.
var errTimeout = errors.New("oxidocs did not finish in time")

// unsupported is one entry of the "unsupported" list oxidocs returns with
// exit code 3: content the engine knows it would not draw.
type unsupported struct {
	Kind  string `json:"kind"`
	Part  string `json:"part"`
	Count int    `json:"count"`
}

// reply is one line oxidocs serve writes on stdout.
type reply struct {
	ID          uint64        `json:"id"`
	OK          bool          `json:"ok"`
	Exit        int           `json:"exit"`
	Error       string        `json:"error"`
	Unsupported []unsupported `json:"unsupported"`
	Pages       int           `json:"pages"`
	Ms          float64       `json:"ms"`
}

// process is one running `oxidocs serve` and the requests waiting on it.
type process struct {
	cmd     *exec.Cmd
	stdin   io.WriteCloser
	pending map[uint64]chan reply // guarded by server.mu
	done    chan struct{}         // closed once the process is reaped
}

// server keeps one `oxidocs serve` process running, so fonts and metrics are
// loaded once and not per document. Requests are JSON lines on its stdin;
// replies come back on stdout in completion order, matched by id. The process
// is started on the first request, and a new one after it dies or is killed.
type server struct {
	binPath string
	threads int
	logger  *slog.Logger

	mu     sync.Mutex
	proc   *process
	nextID uint64
}

func newServer(binPath string, threads int, logger *slog.Logger) *server {
	return &server{binPath: binPath, threads: threads, logger: logger}
}

// start launches a process. Callers hold s.mu.
func (s *server) start() (*process, error) {
	cmd := exec.Command(s.binPath, "serve", "--threads", fmt.Sprint(s.threads))
	stdin, err := cmd.StdinPipe()
	if err != nil {
		return nil, fmt.Errorf("stdin pipe: %w", err)
	}
	stdout, err := cmd.StdoutPipe()
	if err != nil {
		return nil, fmt.Errorf("stdout pipe: %w", err)
	}
	if err := cmd.Start(); err != nil {
		return nil, fmt.Errorf("start %s serve: %w", s.binPath, err)
	}
	p := &process{cmd: cmd, stdin: stdin, pending: make(map[uint64]chan reply), done: make(chan struct{})}
	s.proc = p
	s.logger.Info(fmt.Sprintf("oxidocs serve started (pid %d, %d threads)", cmd.Process.Pid, s.threads))
	go s.read(p, stdout)
	return p, nil
}

// read dispatches p's replies until its stdout closes, then fails whatever is
// still waiting on p.
func (s *server) read(p *process, stdout io.Reader) {
	scanner := bufio.NewScanner(stdout)
	scanner.Buffer(make([]byte, 64*1024), 16*1024*1024)
	for scanner.Scan() {
		var r reply
		if err := json.Unmarshal(scanner.Bytes(), &r); err != nil {
			s.logger.Error(fmt.Sprintf("oxidocs serve: unreadable reply: %s", err))
			continue
		}
		s.mu.Lock()
		ch, ok := p.pending[r.ID]
		delete(p.pending, r.ID)
		s.mu.Unlock()
		if ok {
			ch <- r
		}
	}
	err := p.cmd.Wait()
	s.mu.Lock()
	if s.proc == p {
		s.proc = nil
		s.logger.Warn(fmt.Sprintf("oxidocs serve exited: %v", err))
	}
	for id, ch := range p.pending {
		close(ch)
		delete(p.pending, id)
	}
	s.mu.Unlock()
	close(p.done)
}

// convert asks oxidocs to turn input into output. strict makes it refuse
// (exit 3) documents holding content it would not draw.
func (s *server) convert(ctx context.Context, input, output string, strict bool, timeout time.Duration) (reply, error) {
	s.mu.Lock()
	p := s.proc
	if p == nil {
		var err error
		if p, err = s.start(); err != nil {
			s.mu.Unlock()
			return reply{}, err
		}
	}
	s.nextID++
	id := s.nextID
	ch := make(chan reply, 1)
	p.pending[id] = ch
	line, _ := json.Marshal(map[string]any{"id": id, "input": input, "output": output, "strict": strict})
	_, err := p.stdin.Write(append(line, '\n'))
	s.mu.Unlock()
	if err != nil {
		s.kill(p)
		return reply{}, fmt.Errorf("write request: %w", err)
	}

	timer := time.NewTimer(timeout)
	defer timer.Stop()
	select {
	case r, ok := <-ch:
		if !ok {
			return reply{}, errServerGone
		}
		return r, nil
	case <-timer.C:
		// A conversion that does not finish is stuck in the engine. The
		// process is replaced so it cannot hold a worker thread forever;
		// other requests in flight on it fail with errServerGone and fall
		// back.
		s.kill(p)
		return reply{}, fmt.Errorf("%w (%s)", errTimeout, timeout)
	case <-ctx.Done():
		s.mu.Lock()
		delete(p.pending, id)
		s.mu.Unlock()
		return reply{}, ctx.Err()
	}
}

// kill ends p and makes the next request start a new process, without
// waiting for p to be reaped.
func (s *server) kill(p *process) {
	s.mu.Lock()
	if s.proc == p {
		s.proc = nil
	}
	s.mu.Unlock()
	_ = p.cmd.Process.Kill()
}

// stop closes stdin, which makes oxidocs serve finish its queue and exit.
func (s *server) stop() {
	s.mu.Lock()
	p := s.proc
	s.proc = nil
	s.mu.Unlock()
	if p == nil {
		return
	}
	_ = p.stdin.Close()
	select {
	case <-p.done:
	case <-time.After(5 * time.Second):
		_ = p.cmd.Process.Kill()
	}
}
