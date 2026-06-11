// Browser notification helpers. Everything here is best-effort: the
// Notification constructor itself can throw on some platforms (e.g. Android
// Chrome requires a service worker even when permission is granted), so all
// calls are guarded and failures are silent.

// Ask for permission only when the user starts a scan — unprompted requests
// on page load are penalized by browsers and annoy users.
export function requestNotificationPermission(): void {
  try {
    if ('Notification' in window && Notification.permission === 'default') {
      Notification.requestPermission().catch(() => { /* ignore */ });
    }
  } catch { /* ignore */ }
}

export function showNotification(title: string, body: string): void {
  try {
    if ('Notification' in window && Notification.permission === 'granted') {
      new Notification(title, { body });
    }
  } catch { /* ignore */ }
}
