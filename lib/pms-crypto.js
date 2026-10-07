// Decrypts the PMS Appraisal API's encrypted "Amount" field. The API uses a
// legacy .NET PasswordDeriveBytes (SHA1, 100 iterations) key/IV derivation
// from a fixed password+salt, then AES-256-CBC/PKCS7, with the plaintext
// encoded as UTF-16LE. Verified against PMS's own test vector:
// 'rW1mbGfBg+xeDgp3CZ3UBg==' -> '122980'.

const crypto = require('crypto');

// Supplied by PMS; kept in env (PMS_CRYPTO_PASSWORD / PMS_CRYPTO_SALT), never in source.
const PASSWORD = process.env.PMS_CRYPTO_PASSWORD;
const SALT = process.env.PMS_CRYPTO_SALT;

let cachedKeyIv = null;

function deriveKeyIv() {
  if (cachedKeyIv) return cachedKeyIv;
  if (!PASSWORD || !SALT) throw new Error('PMS_CRYPTO_PASSWORD / PMS_CRYPTO_SALT are not set');
  const pwBytes = Buffer.from(PASSWORD, 'utf8');
  const saltBytes = Buffer.from(SALT, 'ascii');
  let base = crypto.createHash('sha1').update(Buffer.concat([pwBytes, saltBytes])).digest();
  for (let i = 0; i < 98; i++) base = crypto.createHash('sha1').update(base).digest();

  const block1 = crypto.createHash('sha1').update(base).digest();
  const block2 = crypto.createHash('sha1').update(Buffer.concat([Buffer.from([0x31]), base])).digest();
  const block3 = crypto.createHash('sha1').update(Buffer.concat([Buffer.from([0x32]), base])).digest();

  const key = Buffer.concat([block1, block2.subarray(0, 12)]);
  const iv = Buffer.concat([block1.subarray(8, 16), block3.subarray(0, 8)]);
  cachedKeyIv = { key, iv };
  return cachedKeyIv;
}

// Returns the decrypted plaintext string, or null if `cipherB64` is falsy or
// decryption fails (so callers can fall back gracefully rather than crash).
function decryptAmount(cipherB64) {
  if (!cipherB64) return null;
  try {
    const { key, iv } = deriveKeyIv();
    const ciphertext = Buffer.from(cipherB64, 'base64');
    const decipher = crypto.createDecipheriv('aes-256-cbc', key, iv);
    const decrypted = Buffer.concat([decipher.update(ciphertext), decipher.final()]);
    return decrypted.toString('utf16le');
  } catch (err) {
    console.error('Failed to decrypt PMS Amount field:', err.message);
    return null;
  }
}

module.exports = { decryptAmount };
