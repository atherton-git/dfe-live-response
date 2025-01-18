<#
##################################################################################################################
# Win-VssRegistry.ps1
# v1.2
##################################################################################################################

 _       ___           _    __          ____             _      __            
| |     / (_)___      | |  / /_________/ __ \___  ____ _(_)____/ /________  __
| | /| / / / __ \_____| | / / ___/ ___/ /_/ / _ \/ __ / / ___/ __/ ___/ / / /
| |/ |/ / / / / /_____/ |/ (__  |__  ) _, _/  __/ /_/ / (__  ) /_/ /  / /_/ / 
|__/|__/_/_/ /_/      |___/____/____/_/ |_|\___/\__, /_/____/\__/_/   \__, /  
                                               /____/                /____/   

.SYNOPSIS
    This script performs a backup of system artefacts as a ZIP archive, 
    including open files using the Volume Shadow Copy Service (VSS). The
    archive is then moved to the target directory.

.NOTES
    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
    INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A
    PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
    HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
    OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
    SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

# Define source files and directories
$source_files = @(
    "C:\Windows\appcompat\Programs\Amcache.hve"
    "C:\Windows\System32\config\BBI"
    "C:\Windows\System32\config\COMPONENTS"
    "C:\Windows\System32\config\DRIVERS"
    "C:\Windows\System32\config\ELAM"
    "C:\Windows\System32\config\SAM"
    "C:\Windows\System32\config\SECURITY"
    "C:\Windows\System32\config\SOFTWARE"
    "C:\Windows\System32\config\SYSTEM"
    "C:\Windows\System32\SRUDB.dat"
)
$source_dirs = @(
	"C:\Windows\System32\winevt\Logs"
	"C:\Windows\Prefetch"
	"C:\Windows\System32\SUM"
)

# Get the list of user accounts as an array
$user_accounts = @(Get-ChildItem C:\Users -Name)

# Append each user's NTUSER.dat path to $source_files
foreach ($user in $user_accounts) {
    $source_files += @{
        Path = "C:\Users\$user\NTUSER.dat"
        Renamed = "$user`_NTUSER.dat"
    }
}

###############################################################################
# Function: Write-Log
###############################################################################
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "Info"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Write-Host $logEntry
}

$target_dir = "$env:systemdrive\nmc-edr-lr\vss-export"
$date = Get-Date -Format yyyy-MM-dd

# Ensure target directory exists
if (-Not (Test-Path -Path $target_dir)) {
    Write-Log "Target directory does not exist. Creating: $target_dir" -Level "Info"
    New-Item -Path $target_dir -ItemType Directory -Force | Out-Null
}

$temp_shadow_link = "C:\shadowcopy_$date"
$unix_time = Get-Date -UFormat %s -Millisecond 0
$archive_filename = "$date_$unix_time.zip"
$temp_archive_full_path = Join-Path -Path $env:TEMP -ChildPath $archive_filename
$temp_collected_dir = Join-Path -Path $env:TEMP -ChildPath "CollectedFiles_$date"

# Create a temporary directory for collected files
if (-Not (Test-Path -Path $temp_collected_dir)) {
    New-Item -Path $temp_collected_dir -ItemType Directory -Force | Out-Null
}

Write-Log "Creating new shadow copy snapshot." -Level "Info"
$s1 = (Get-WmiObject -List Win32_ShadowCopy).Create("C:\", "ClientAccessible")
$s2 = Get-WmiObject Win32_ShadowCopy | Where-Object { $_.ID -eq $s1.ShadowID }
$d  = $s2.DeviceObject + "\"

if (Test-Path -Path $temp_shadow_link) {
    Write-Log "Temporary shadow link exists. Deleting: $temp_shadow_link" -Level "Warning"
    Remove-Item -Path $temp_shadow_link -Recurse -Force
}
cmd /c mklink /d $temp_shadow_link $d

# Collect files and directories into the temporary directory
foreach ($file in $source_files) {
    if ($file -is [Hashtable]) {
        # Handle NTUSER.dat files with renaming
        $shadow_file = $file.Path -replace "^C:", $temp_shadow_link
        $renamed_file = Join-Path -Path $temp_collected_dir -ChildPath $file.Renamed

        if (Test-Path -Path $shadow_file) {
            Write-Log "Copying and renaming NTUSER.dat for user: $shadow_file -> $renamed_file" -Level "Info"
            Copy-Item -Path $shadow_file -Destination $renamed_file -Force
        } else {
            Write-Log "NTUSER.dat file not found in shadow copy for user: $shadow_file" -Level "Warning"
        }
    } else {
        # Handle regular files
        $shadow_file = $file -replace "^C:", $temp_shadow_link
        if (Test-Path -Path $shadow_file) {
            Write-Log "Copying file to temporary directory: $shadow_file" -Level "Info"
            Copy-Item -Path $shadow_file -Destination $temp_collected_dir -Force
        } else {
            Write-Log "File not found in shadow copy: $shadow_file" -Level "Warning"
        }
    }
}

foreach ($dir in $source_dirs) {
    $shadow_dir = $dir -replace "^C:", $temp_shadow_link
    if (Test-Path -Path $shadow_dir) {
        Write-Log "Copying directory to temporary directory: $shadow_dir" -Level "Info"
        Copy-Item -Path $shadow_dir -Destination $temp_collected_dir -Recurse -Force
    } else {
        Write-Log "Directory not found in shadow copy: $shadow_dir" -Level "Warning"
    }
}

# Create the ZIP archive using System.IO.Compression.ZipFile
try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    Write-Log "Creating ZIP archive: $temp_archive_full_path" -Level "Info"
    [System.IO.Compression.ZipFile]::CreateFromDirectory($temp_collected_dir, $temp_archive_full_path)
} catch {
    Write-Log "Failed to create ZIP archive. Error: $_" -Level "Error"
}

# Clean up shadow snapshot, temporary link, and collected files
Write-Log "Deleting shadow snapshot and cleaning up temporary link." -Level "Info"
$s2.Delete()
cmd /c rmdir $temp_shadow_link

Write-Log "Cleaning up temporary collected files directory." -Level "Info"
Remove-Item -Path $temp_collected_dir -Recurse -Force

# Move the archive to the target directory
Write-Log "Moving archive to target directory: $target_dir" -Level "Info"
robocopy $env:TEMP $target_dir $archive_filename /MOVE /R:3 /np

Write-Log "Backup completed successfully at $(Get-Date -Format yyyy-MM-dd_HH:mm:ss)." -Level "Info"
