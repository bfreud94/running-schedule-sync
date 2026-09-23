param(
    [Parameter(Position = 0)]
    [string]$Since
)

$ErrorActionPreference = 'Stop'
$arguments = @('local/sync-activity-mocks.js')
if ($Since) { $arguments += $Since }
node @arguments