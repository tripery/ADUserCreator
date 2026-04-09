# Project Structure

## Main source folders

- `src/`
  Desktop application entrypoint and app-local folders.
- `src/main.ps1`
  Desktop entrypoint.
- `src/ui/`
  WinForms / desktop UI scripts.
- `src/core/`
  Shared business logic used by both desktop and web backends.
- `src/core/ad/`
  Active Directory logic, naming, transliteration, preview/create helpers.
- `src/core/common/`
  Shared helpers such as logging and password generation.
- `src/core/excel/`
  Excel import and parsing logic for desktop workflows.

- `webapi/`
  PowerShell HTTP API used by the React frontend.

- `webui-react/`
  Main web frontend built with React + Vite.
- `webui-react/src/`
  React application source.

- `legacy/webui/`
  Archived static web prototype kept only for reference.

- `build/`
  Packaging and installer scripts.

- `assets/`
  Shared static assets for packaging and desktop app branding.

## Convenience scripts

- `scripts/start.ps1`
  Main entry point for local development and Docker startup.
- `scripts/start-dev.ps1`
  Wrapper for `scripts/start.ps1 -UiMode Local`.
- `scripts/start-all.ps1`
  Wrapper for `scripts/start.ps1 -UiMode Docker`.
- `scripts/compose-up.cmd`
  CMD wrapper for `scripts/start.ps1`.

## Generated / build output

These should not be treated as source folders:

- `dist/`
  Desktop build output.
- `release/`
  Release artifacts.
- `ADUserCreator/`
  Packaged output snapshot.
- `webui-react/dist/`
  React production build.
- `webui-react/node_modules/`
  Installed frontend dependencies.

## Practical rule

If you are changing application logic, work mainly in:

- `src/core/`
- `src/ui/`
- `src/main.ps1`
- `webapi/`
- `webui-react/src/`
- `build/`

Avoid editing generated folders unless you are verifying build output.
Avoid editing `legacy/` unless you intentionally need the old prototype.
