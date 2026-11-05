import * as React from 'react';
import {
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
} from '@fluentui/react-components';

export interface ConfirmOptions {
  title: string;
  body: string;
  confirmLabel: string;
  cancelLabel: string;
  onConfirm: () => void;
}

// Themed, keyboard- and screen-reader-friendly replacement for window.confirm.
// Usage: const { confirm, dialog } = useConfirmDialog(); render {dialog} once, and call
// confirm({...}) from an event handler — onConfirm runs only if the user accepts.
export function useConfirmDialog(): { confirm: (options: ConfirmOptions) => void; dialog: React.ReactElement } {
  const [options, setOptions] = React.useState<ConfirmOptions | null>(null);

  const close = (): void => setOptions(null);

  const dialog = (
    <Dialog
      open={options !== null}
      modalType="alert"
      onOpenChange={(_, data) => { if (!data.open) close(); }}
    >
      <DialogSurface>
        <DialogBody>
          <DialogTitle>{options?.title}</DialogTitle>
          <DialogContent>{options?.body}</DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={close}>{options?.cancelLabel}</Button>
            <Button
              appearance="primary"
              onClick={() => {
                const action = options?.onConfirm;
                close();
                action?.();
              }}
            >
              {options?.confirmLabel}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );

  return { confirm: setOptions, dialog };
}
