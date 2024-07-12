import CryptoJS from "crypto-js";

const encryptedFreepikApiKey =
  "U2FsdGVkX18258oINhW6ItarhxVnw+paVm8IdfMpDwfw8+I+aJeBKEBK6dz/6wFptAVG5SB+N3ljsBaqN9X2yw==";

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
