function Invoke-SPOVTRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [scriptblock]$Action,

        [ValidateRange(1, 20)]
        [int]$MaxRetries = 8,

        [ValidateRange(2, 120)]
        [int]$MaxBackoffSeconds = 60
    )

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            return & $Action
        }
        catch {
            $msg = $_.Exception.Message
            $inner = $_.Exception.InnerException?.Message

            $isTransient =
                ($msg -match '(?i)throttl|429|503|temporarily unavailable|timeout|server busy') -or
                ($inner -match '(?i)throttl|429|503|temporarily unavailable|timeout|server busy')

            if (-not $isTransient -or $attempt -eq $MaxRetries) {
                throw
            }

            $sleep = [Math]::Min($MaxBackoffSeconds, [Math]::Max(2, [Math]::Pow(2, $attempt)))
            Start-Sleep -Seconds $sleep
        }
    }
}