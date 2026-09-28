// Small helper for localized strings that contain positional placeholders,
// e.g. "Check access for {0}". Used instead of template literals so the
// English source string (and every translated locale file) can freely
// reorder the placeholder for that language's grammar.
export function formatString(template: string, ...args: (string | number)[]): string {
  return template.replace(/\{(\d+)\}/g, (match, index: string) => {
    const i = parseInt(index, 10);
    return i < args.length ? String(args[i]) : match;
  });
}
