# MetoMail (outlook-local-ai-chat)

Windows Outlook COM add-in (.NET Framework 4.8, C# 7.3, classic csproj).
It cannot be built or run on Linux — the Windows CI workflow
(`.github/workflows/build.yml`) is the compile/test gate.

## Git workflow

- This is a personal dev repository. **Always commit directly to `main` and
  push immediately after each change set.** The user pulls `main` on a work
  machine to test. Do not create side branches or pull requests unless
  explicitly asked.
- Every push to `main` triggers CI: MSBuild, guardrail tests
  (`tests/GuardrailTests`), the static capability scan
  (`scripts/Test-Guardrails.ps1`), and republishing the installer to the
  `continuous` release that the README download link points at.

## Code conventions

- C# 7.3 only (no target-typed `new`, ranges, or switch expressions).
- String concatenation over interpolation; match the existing wrapping style.
- New source files must be added to `OutlookLocalAIChat.csproj` (classic
  csproj — no globbing).
- Security boundaries are load-bearing: the static scan asserts exact strings
  in several files (tool names, draft authorization, working-set caps).
  Check `scripts/Test-Guardrails.ps1` before renaming or rewording anything
  it references.
- Guardrail tests use only public APIs (no InternalsVisibleTo) and a
  hand-rolled runner in `tests/GuardrailTests/Program.cs` — register new
  tests in `Main`.
