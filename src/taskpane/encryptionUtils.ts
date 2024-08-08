import CryptoJS from "crypto-js";

const encryptedFreepikApiKey =
  "U2FsdGVkX1+uakX7Pa7rDwZUW41gsxUcc5Mk1Zf9Ff/RkAhy5TT6snpzfSoe7kn+A7ulqhvrLMasq4hJ/KkwoA==";

export function storeFreepikApiKeySecret() {
  document.getElementById("save-api-key-secret").onclick = () => {
    const encryptionSecret = (<HTMLInputElement>document.getElementById("api-key-secret")).value;
    localStorage.setItem("apiKeySecret", encryptionSecret);
    (<HTMLTextAreaElement>document.getElementById("api-key-secret")).value = "API key secret stored";
  };
}

export function getDecryptedFreepikApiKey(): string {
  const decryptedBytes = CryptoJS.AES.decrypt(encryptedFreepikApiKey, localStorage.getItem("apiKeySecret"));
  return decryptedBytes.toString(CryptoJS.enc.Utf8);
}
