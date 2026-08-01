# SplitVision agent

Use this workspace as a Windows-focused PDF processing application. Prioritize compatibility with Tkinter, PyPDF2, pypdfium2, winocr, winrt, and windnd.

## Expectations
- Keep the app fully in-process and avoid external OCR or PDF tools.
- Preserve the current GUI behavior and the output log format.
- Favor small, targeted changes over large rewrites.
- Prefer validation with py_compile or a local run when possible.

## Important paths
- Main app: splitvision.py
- Build outputs: splitvision.build/, splitvision.dist/, splitvision.onefile-build/

## Notes
The current implementation already uses Windows native OCR and pypdfium2 for rendering. New work should build on that approach rather than reintroducing subprocess-based dependencies.
