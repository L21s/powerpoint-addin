export async function getImageAsBase64(url: string): Promise<string> {
  const blob = await fetch(url).then(r => r.blob());
  return new Promise((res, rej) => {
    const reader = new FileReader();
    reader.onload = () => res(reader.result as string);
    reader.onerror = () => rej(reader.error);
    reader.readAsDataURL(blob);
  });
}
