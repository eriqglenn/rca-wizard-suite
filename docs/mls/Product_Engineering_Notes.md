# Product and Engineering Notes

v5 adds the applicant complete/sign-posting letter as the second document module.

## UX decisions

- Keep Step 1 GIS labels unchanged.
- Keep Step 2 official RCA calculation layout unchanged.
- Store user-entered case fields on `Case Info`.
- Keep map selection in the wizard UI.
- Keep preview/generate actions as buttons, not workbook fields.

## Applicant address handling

`Applicant_Address_2` is optional. The generated letter uses `Applicant_Address_2_Block`, which becomes a new line only when the field has a value.

## Document modules

- Notification mail-out: one preview or all valid owner rows; map inserted after each letter.
- Applicant sign letter: single letter from `Case Info`; no map required.
