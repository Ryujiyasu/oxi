// Asks, once, what a person makes of Oxi — after they have actually used it.
//
// Nothing here runs unless the build was given somewhere to send to, and
// nothing leaves the machine until the person has read what is about to go and
// pressed the button. Closing the card sends nothing and asks again later;
// "don't ask again" sends nothing and never asks. The card itself makes no
// request: it is not a beacon wearing a question.
//
// It waits for real use rather than for time. Five saved documents is someone
// who has done work in it, which is the only person whose answer is worth
// having.

const ENDPOINT = '{{FEEDBACK_ENDPOINT}}';
const AFTER = 5;
const SAVES = 'oxi-saves';
const ASKED = 'oxi-rated';

const WORDS = {
  en: {
    ask: 'How is Oxi working out?',
    hint: 'Five documents in. Worth a minute?',
    placeholder: 'Anything you would change (optional)',
    sends: 'What will be sent:',
    send: 'Send',
    later: 'Later',
    never: "Don't ask again",
    thanks: 'Thank you.',
    failed: 'It could not be sent. Nothing was kept.',
    rating: (n) => `Rating ${n} of 5`,
  },
  ja: {
    ask: 'Oxi の使い心地はいかがですか',
    hint: '5 つ保存されました。少しだけ聞かせてください',
    placeholder: '直してほしいところがあれば（任意）',
    sends: '送られる内容:',
    send: '送信',
    later: '後で',
    never: '今後表示しない',
    thanks: 'ありがとうございます。',
    failed: '送信できませんでした。何も残っていません。',
    rating: (n) => `評価 5 段階中 ${n}`,
  },
};

function held(key, fallback) {
  try {
    const value = localStorage.getItem(key);
    return value === null ? fallback : value;
  } catch {
    return fallback;
  }
}
function keep(key, value) {
  try { localStorage.setItem(key, String(value)); } catch { /* a private window */ }
}

// The editors write this when someone picks a language, so the card speaks
// whichever one they picked rather than guessing again.
const lang = held('oxi-lang', (navigator.language || 'en').startsWith('ja') ? 'ja' : 'en');
const say = (key) => (WORDS[lang] || WORDS.en)[key];

function style() {
  const sheet = document.createElement('style');
  sheet.textContent = `
    #oxi-rate {
      position: fixed; right: 18px; bottom: 18px; z-index: 2147483000;
      width: 320px; max-width: calc(100vw - 36px);
      background: #fff; color: #1f1f1f;
      border: 1px solid #d8d8d8; border-radius: 10px;
      box-shadow: 0 8px 28px rgba(0,0,0,.16);
      font: 13px/1.55 'Segoe UI', -apple-system, 'Hiragino Kaku Gothic ProN', Meiryo, sans-serif;
      padding: 16px; display: flex; flex-direction: column; gap: 10px;
    }
    #oxi-rate h3 { margin: 0; font-size: 14px; font-weight: 600; }
    #oxi-rate p { margin: 0; color: #6b6b6b; font-size: 12px; }
    #oxi-rate .stars { display: flex; gap: 4px; }
    #oxi-rate .stars button {
      font: inherit; font-size: 20px; line-height: 1;
      background: none; border: 0; padding: 2px 3px; cursor: pointer;
      color: #c9c9c9;
    }
    #oxi-rate .stars button.on { color: #E8A33D; }
    #oxi-rate textarea {
      font: inherit; font-size: 12px; resize: vertical; min-height: 56px;
      border: 1px solid #d8d8d8; border-radius: 5px; padding: 7px 8px;
      color: inherit; background: #fff;
    }
    #oxi-rate .payload {
      font-family: ui-monospace, Consolas, monospace; font-size: 11px;
      color: #6b6b6b; background: #f6f6f4; border-radius: 5px;
      padding: 7px 8px; white-space: pre-wrap; word-break: break-all;
      max-height: 86px; overflow: auto;
    }
    #oxi-rate .ends { display: flex; align-items: center; gap: 8px; }
    #oxi-rate .ends .gap { flex: 1; }
    #oxi-rate .ends button {
      font: inherit; font-size: 12px; padding: 6px 12px; border-radius: 5px;
      border: 1px solid #d8d8d8; background: #f6f6f4; color: #1f1f1f; cursor: pointer;
    }
    #oxi-rate .ends button.go { background: #C7462F; border-color: #C7462F; color: #fff; }
    #oxi-rate .ends button:disabled { opacity: .5; cursor: default; }
    #oxi-rate .ends .plain { border-color: transparent; background: none; color: #6b6b6b; }
  `;
  document.head.appendChild(sheet);
}

function show() {
  style();
  const card = document.createElement('div');
  card.id = 'oxi-rate';
  card.innerHTML = `
    <h3></h3>
    <p class="hint"></p>
    <div class="stars"></div>
    <textarea hidden></textarea>
    <div class="sends" hidden>
      <p style="margin-bottom:4px"></p>
      <div class="payload"></div>
    </div>
    <div class="ends">
      <button class="plain never"></button>
      <span class="gap"></span>
      <button class="later"></button>
      <button class="go" disabled></button>
    </div>`;
  document.body.appendChild(card);

  card.querySelector('h3').textContent = say('ask');
  card.querySelector('.hint').textContent = say('hint');
  card.querySelector('textarea').placeholder = say('placeholder');
  card.querySelector('.sends p').textContent = say('sends');
  card.querySelector('.never').textContent = say('never');
  card.querySelector('.later').textContent = say('later');
  card.querySelector('.go').textContent = say('send');

  const stars = card.querySelector('.stars');
  const note = card.querySelector('textarea');
  const sends = card.querySelector('.sends');
  const payload = card.querySelector('.payload');
  const go = card.querySelector('.go');
  let picked = 0;

  /// Exactly what the button will send, spelled out. Nothing about the
  /// document is in it, and there is nothing here that is not shown.
  function parcel() {
    return {
      rating: picked,
      comment: note.value.trim(),
      version: document.querySelector('meta[name="oxi-version"]')?.content || '',
      platform: navigator.platform || '',
    };
  }
  function restate() {
    payload.textContent = JSON.stringify(parcel(), null, 1);
  }

  for (let n = 1; n <= 5; n += 1) {
    const star = document.createElement('button');
    star.type = 'button';
    star.textContent = '★';
    star.title = say('rating')(n);
    star.onclick = () => {
      picked = n;
      for (const [at, one] of [...stars.children].entries()) {
        one.classList.toggle('on', at < n);
      }
      note.hidden = false;
      sends.hidden = false;
      go.disabled = false;
      restate();
    };
    stars.appendChild(star);
  }
  note.addEventListener('input', restate);

  const close = () => card.remove();
  card.querySelector('.later').onclick = close;
  card.querySelector('.never').onclick = () => { keep(ASKED, 'never'); close(); };
  go.onclick = async () => {
    go.disabled = true;
    try {
      const reply = await fetch(ENDPOINT, {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify(parcel()),
      });
      if (!reply.ok) throw new Error(String(reply.status));
      keep(ASKED, 'sent');
      card.querySelector('h3').textContent = say('thanks');
      card.querySelector('.hint').remove();
      stars.remove(); note.remove(); sends.remove();
      card.querySelector('.ends').remove();
      setTimeout(close, 2200);
    } catch {
      card.querySelector('.hint').textContent = say('failed');
      go.disabled = false;
    }
  };
}

/// Counted here rather than in each editor, so the three of them agree on what
/// "used it" means without holding a count of their own.
window.addEventListener('oxi:saved', () => {
  if (held(ASKED, '')) return;
  const now = Number(held(SAVES, '0')) + 1;
  keep(SAVES, now);
  if (now === AFTER && !document.getElementById('oxi-rate')) show();
});
