export function storeFreepikApiKeyEncryptionSecret() {
  document.getElementById("save-encryption-secret").onclick = () => {
    const encryptionSecret = (<HTMLInputElement>document.getElementById("encryption-secret")).value;
    localStorage.setItem("apiKeyEncryptionSecret", encryptionSecret);
    (<HTMLTextAreaElement>document.getElementById("encryption-secret")).value = "Secret stored";
  };
}
