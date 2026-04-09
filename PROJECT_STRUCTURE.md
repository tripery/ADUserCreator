# Project Structure

## Main source folders

- `src/`
  PowerShell desktop application source code.
- `src/ad/`
  Active Directory logic, naming, transliteration, provisioning helpers.
- `src/common/`
  Shared helpers such as logging and password generation.
- `src/excel/`
  Excel import and parsing logic.
- `src/ui/`
  WinForms / desktop UI scripts.

- `build/`
  Packaging and installer scripts.

- `assets/`
  Shared static assets for packaging and desktop app branding.

## Generated / build output

These should not be treated as source folders:

- `dist/`
  Desktop build output.
- `release/`
  Release artifacts.
- `ADUserCreator/`
  Packaged output snapshot.

## Practical rule

If you are changing application logic, work mainly in:

- `src/`
- `build/`

Avoid editing generated folders unless you are verifying build output.
