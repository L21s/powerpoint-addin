import CryptoJS from "crypto-js";

const encryptedFreepikApiKey =
  "U2FsdGVkX1+uakX7Pa7rDwZUW41gsxUcc5Mk1Zf9Ff/RkAhy5TT6snpzfSoe7kn+A7ulqhvrLMasq4hJ/KkwoA==";

export function storeEncryptionKey() {
  document.getElementById("save-encryption-key").onclick = () => {
    const encryptionSecret = (<HTMLInputElement>document.getElementById("encryption-key")).value;
    localStorage.setItem("encryptionKey", encryptionSecret);
    (<HTMLTextAreaElement>document.getElementById("encryption-key")).value = "Encryption key saved";
  };
}

export function getDecryptedFreepikApiKey(): string {
  const decryptedBytes = CryptoJS.AES.decrypt(encryptedFreepikApiKey, localStorage.getItem("encryptionKey"));
  return decryptedBytes.toString(CryptoJS.enc.Utf8);
}
