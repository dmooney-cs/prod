# --------------------------
# SHA-256 Verification (Hash URL)
# --------------------------
if (-not $SkipHashCheck) {

    $expected = Get-ExpectedSha256 -HashUri $HashUrl
    if ([string]::IsNullOrWhiteSpace($expected)) {
        Write-Host "⚠️ Could not retrieve/parse expected SHA-256 from hash URL. Aborting to be safe." -ForegroundColor Yellow
        Write-Host ("   Hash URL: {0}" -f $HashUrl) -ForegroundColor Yellow
        try { Remove-Item -LiteralPath $ZipPath -Force -ErrorAction SilentlyContinue } catch { }
        return
    }

    try {
        $actual = (Get-FileHash -Algorithm SHA256 -Path $ZipPath).Hash.ToUpper()
        if ($actual -ne $expected) {
            Write-Host "❌ Hash mismatch! ZIP will be discarded." -ForegroundColor Red
            Write-Host ("   Expected: {0}" -f $expected) -ForegroundColor Red
            Write-Host ("   Actual  : {0}" -f $actual) -ForegroundColor Red
            try { Remove-Item -LiteralPath $ZipPath -Force -ErrorAction SilentlyContinue } catch { }
            return
        } else {
            Write-Host "✅ File hash verified (SHA-256)." -ForegroundColor Green
        }
    } catch {
        Write-Host ("❌ ERROR computing SHA-256: {0}" -f $_.Exception.Message) -ForegroundColor Red
        try { Remove-Item -LiteralPath $ZipPath -Force -ErrorAction SilentlyContinue } catch { }
        return
    }

} else {
    Write-Host "⚠️ Hash verification skipped by user switch." -ForegroundColor Yellow
}
