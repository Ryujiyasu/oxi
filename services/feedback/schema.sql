-- What a person chose to send, and nothing that would identify them.
create table if not exists feedback (
  id integer primary key autoincrement,
  at text not null,
  rating integer not null,
  comment text,
  version text,
  platform text
);
create index if not exists feedback_at on feedback (at);

-- One row per address per day, holding only how many times it has sent. Rows
-- older than a couple of days are of no use; sweep them when it suits:
--   delete from senders where key < date('now', '-2 days');
create table if not exists senders (
  key text primary key,
  times integer not null
);
