// Reads the tax team's reply to the TDS email: an amount to deduct, or NIL.
// Deliberately conservative — anything ambiguous is reported as 'unclear' so a
// person reviews it, rather than guessing a deduction amount.

const MAX_CHARS = 600; // the answer sits at the top; skip signatures further down

function parseTdsReply(text) {
  const head = String(text || '').replace(/\r/g, '').trim().slice(0, MAX_CHARS)
    .replace(/[\w.+-]+@[\w-]+(\.[\w-]+)+/g, ' ')            // email addresses
    .replace(/\b\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}\b/g, ' ');  // dates like 07/10/2026

  const isNil = /\b(nil|zero|no tds|not applicable|n\/a)\b/i.test(head);

  const marked = [...head.matchAll(/(?:₹|rs\.?|inr)\s*(\d[\d,]*(?:\.\d+)?)|(\d[\d,]*(?:\.\d+)?)\s*\/-/gi)]
    .map((m) => m[1] || m[2]);
  // Bare numbers only count when there's no NIL and no currency-marked amount;
  // 9+ digit runs are phone numbers in a signature, never a TDS amount.
  const bare = [...head.matchAll(/\b\d[\d,]*(?:\.\d+)?\b/g)].map((m) => m[0]).filter((n) => n.replace(/\D/g, '').length <= 8);
  const candidates = marked.length ? marked : (!isNil && bare.length === 1 ? bare : []);
  const amounts = [...new Set(candidates.map((n) => Number(n.replace(/,/g, ''))))].filter((n) => n > 0);

  if (isNil && amounts.length) return { kind: 'unclear' };
  if (isNil) return { kind: 'nil' };
  if (amounts.length === 1) return { kind: 'amount', amount: amounts[0] };
  return { kind: 'unclear' };
}

module.exports = { parseTdsReply };
