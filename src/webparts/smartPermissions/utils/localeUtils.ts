let displayLocale: string | undefined;

/** Sets the locale used for dates and numbers — the SharePoint page language, not the browser's. */
export function setDisplayLocale(locale: string | undefined): void {
  displayLocale = locale || undefined;
}

export function formatDateTime(value: string | number | Date): string {
  const date = new Date(value);
  try {
    return date.toLocaleString(displayLocale);
  } catch {
    return date.toLocaleString(); // unsupported locale tag
  }
}

export function formatNumber(value: number): string {
  try {
    return value.toLocaleString(displayLocale);
  } catch {
    return value.toLocaleString();
  }
}

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
