# Interface design

## Register

Product interface. The tool is used at a desk, usually in a well-lit church office, by staff concentrating on detailed service preparation.

## Visual system

- System sans-serif for product controls and data; Georgia appears only in the sign-in brand statement and circular TWPC mark.
- Restrained palette: white working surfaces, blue-grey navigation and canvas layers, TWPC burgundy for primary actions and current selection.
- Six-pixel controls, ten-pixel working surfaces, fourteen-pixel dialogs. Borders define structure; small shadows are reserved for floating or focused layers.
- The desktop builder has an order panel, contextual editor, and readiness panel. At narrower desktop widths the readiness panel collapses before the editing surface does.

## Interaction

- Autosave begins after 900 ms of inactivity and always exposes a visible save state.
- Dragging and Alt+Arrow keyboard movement provide equivalent ordering controls.
- Warnings remain non-blocking. Missing authentication, CSRF, or the editing lease blocks mutation with a clear reason.
- Motion communicates save, progress, dialog, and toast state in 120–180 ms. Reduced-motion preferences remove meaningful movement.
