import readline from 'readline';
import cryptoJS from 'crypto-js';

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
});

rl.question("Enter Freepik API key from 1Password: ", (apiKey) => {
  rl.question("Enter encryption key from 1Password: ", (secret) => {
    console.log(`Encrypted API key: ${cryptoJS.AES.encrypt(apiKey, secret).toString()}`);
    rl.close();
  });
});
