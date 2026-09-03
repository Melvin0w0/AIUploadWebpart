export function locFormat(template: string | undefined, fallback: string, ...values: string[]): string {
  let text = template || fallback;
  for (let i = 0; i < values.length; i++) {
    text = text.split(`{${i}}`).join(values[i]);
  }
  return text;
}
