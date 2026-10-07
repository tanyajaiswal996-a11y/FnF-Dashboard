// Decides whether an employee's emailed reply clearly accepts/acknowledges the
// FnF. Deliberately strict: queries, disagreements, conditions, thanks-only or
// "noted"-only replies are NOT an acknowledgement. When in doubt: 'not_ack'.

const MAX_CHARS = 500; // the answer is at the top of the reply; skip signatures

const STRONG_POSITIVE = /\b(acknowledg(?:e|ed|es|ement|ment)|accept(?:ed|s)?|agree(?:d|s)?|confirm(?:ed|s)?|approve(?:d|s)?|no objections?|(?:ok|okay|fine|good) with)\b/i;

const NEGATIVE_OR_QUERY = /\?|\b(not|no|never|cannot|can't|cant|won't|don't|dont|doesn't|didn't|disagree\w*|dispute\w*|incorrect|wrong|mistake\w*|error|issue\w*|concern\w*|quer(?:y|ies)|clarif\w*|doubt\w*|why|how|explain|explanation|please (?:check|review|revise|recheck|re-check|correct|share|send|confirm)|kindly (?:check|review|revise|recheck|correct|explain|share)|but|however|although|unless|until|if|once|after|when|provided|subject to|only if|on condition|unable|reject\w*|declin\w*|object\w*|discrepanc\w*|less|short|missing|revis\w*|recalculat\w*|pending|waiting)\b/i;

// Drops the quoted earlier thread ("On ... wrote:", "From:", > quotes, Outlook separators).
function stripQuoted(text) {
  const t = String(text || '').replace(/\r/g, '');
  const cut = t.search(/^(>|On .{0,120}wrote:|From:|-{2,}\s*Original Message|_{5,})/m);
  return cut === -1 ? t : t.slice(0, cut);
}

function classifyAck(text) {
  let head = stripQuoted(text).trim().slice(0, MAX_CHARS);
  head = head
    .replace(/[\w.+-]+@[\w-]+(\.[\w-]+)+/g, ' ')   // email addresses
    .replace(/\bno objections?\b/gi, 'noobjection') // "no objection" is positive, not a "no"
    .replace(/\s+/g, ' ');
  if (!head) return 'not_ack';
  if (NEGATIVE_OR_QUERY.test(head)) return 'not_ack';
  return STRONG_POSITIVE.test(head.replace(/noobjection/gi, 'no objection')) ? 'ack' : 'not_ack';
}

module.exports = { classifyAck };
