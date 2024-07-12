export function storeFreepikApiKeyEncryptionSecret() {
  document.getElementById("save-api-key-secret").onclick = () => {
    const encryptionSecret = (<HTMLInputElement>document.getElementById("api-key-secret")).value;
    localStorage.setItem("apiKeySecret", encryptionSecret);
    (<HTMLTextAreaElement>document.getElementById("api-key-secret")).value = "API key secret stored";
  };
}
