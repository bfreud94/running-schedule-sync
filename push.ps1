$ErrorActionPreference = 'Stop'

git push
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

clasp push
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}