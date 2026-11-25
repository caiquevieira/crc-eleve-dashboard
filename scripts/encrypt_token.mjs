// Gera site/auth.json: o token do GitHub criptografado com login + senha do admin.
// O login e a senha NÃO são gravados em lugar nenhum; sem eles o arquivo não abre.
//
// Uso (o token fica em token.txt, que é apagado ao final):
//   node scripts/encrypt_token.mjs <login> <senha> [token.txt]
import { pbkdf2Sync, randomBytes, createCipheriv } from 'node:crypto';
import { readFileSync, writeFileSync, unlinkSync, existsSync } from 'node:fs';

const [login, senha, tokenFile = 'token.txt'] = process.argv.slice(2);
if (!login || !senha) {
  console.error('Uso: node scripts/encrypt_token.mjs <login> <senha> [token.txt]');
  process.exit(1);
}
if (!existsSync(tokenFile)) {
  console.error(`Arquivo ${tokenFile} não encontrado. Cole o token do GitHub nele e rode de novo.`);
  process.exit(1);
}
const token = readFileSync(tokenFile, 'utf8').trim();
if (!/^(github_pat_|ghp_)/.test(token)) {
  console.error('O conteúdo de token.txt não parece um token do GitHub (github_pat_…).');
  process.exit(1);
}

const ITER = 600000;
const salt = randomBytes(16);
const iv = randomBytes(12);
const key = pbkdf2Sync(`${login}\n${senha}`, salt, ITER, 32, 'sha256');
const cipher = createCipheriv('aes-256-gcm', key, iv);
const ct = Buffer.concat([cipher.update(token, 'utf8'), cipher.final(), cipher.getAuthTag()]);

const b64 = (b) => b.toString('base64');
writeFileSync('site/auth.json', JSON.stringify({ v: 1, iter: ITER, salt: b64(salt), iv: b64(iv), ct: b64(ct) }) + '\n');
unlinkSync(tokenFile);
console.log('site/auth.json gerado e token.txt apagado.');
