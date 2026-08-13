# Contributing

1. Do not commit Microsoft payloads or generated binaries.
2. Keep the stock Teams COM identity untouched.
3. Build with `scripts/build/Build-Windows.ps1` and run
   `scripts/test/Test-TmaInstallation.ps1` before proposing a change.
4. Test Outlook x64 with and without the official Teams add-in enabled.
5. Never log tokens, account IDs or meeting URLs.
