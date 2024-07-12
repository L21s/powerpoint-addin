const readline = require("readline");
const cryptoJS = require("crypto-js");

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
});

rl.question("Enter Freepik API key from 1Password: ", (apiKey) => {
  rl.question("Enter API key secret from 1Password: ", (secret) => {
    console.log(`Encrypted API key: ${cryptoJS.AES.encrypt(apiKey, secret).toString()}`);
    rl.close();
  });
});
