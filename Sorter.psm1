Set-StrictMode -Version Latest
Add-Type -AssemblyName Microsoft.VisualBasic

$script:CategoryDefinitions = @(
    @{
        Name = 'Images'
        Extensions = @(
            '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tif', '.tiff', '.webp', '.svg',
            '.heic', '.ico', '.raw'
        )
        NameSuffixes = @()
    },
    @{
        Name = 'Video'
        Extensions = @(
            '.mp4', '.mov', '.avi', '.mkv', '.wmv', '.webm', '.m4v', '.mpg', '.mpeg', '.ts'
        )
        NameSuffixes = @()
    },
    @{
        Name = 'Project Files'
        Extensions = @(
            '.psd', '.psb', '.veg', '.vegtemplate', '.vf', '.prproj', '.aep', '.aup3',
            '.blend', '.afphoto', '.afdesign', '.afpub', '.clip', '.kra', '.xcf'
        )
        NameSuffixes = @(
            '.veg.bak'
        )
    },
    @{
        Name = 'Documents'
        Extensions = @(
            '.txt', '.rtf', '.md', '.doc', '.docx', '.odt', '.pdf', '.csv', '.tsv',
            '.xls', '.xlsx', '.ods', '.ppt', '.pptx', '.odp'
        )
        NameSuffixes = @()
    }
)

function Get-ExtensionFolderName {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Extension
    )

    if ([string]::IsNullOrWhiteSpace($Extension)) {
        return 'No Extension'
    }

    return $Extension.TrimStart('.').ToUpperInvariant()
}

function Get-CategoryFolderName {
    param(
        [Parameter(Mandatory)]
        [string]$FileName,

        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Extension
    )

    $normalizedName = $FileName.ToLowerInvariant()
    $normalizedExtension = $Extension.ToLowerInvariant()

    foreach ($definition in $script:CategoryDefinitions) {
        foreach ($suffix in $definition.NameSuffixes) {
            if ($normalizedName.EndsWith($suffix, [System.StringComparison]::OrdinalIgnoreCase)) {
                return $definition.Name
            }
        }

        if ($definition.Extensions -contains $normalizedExtension) {
            return $definition.Name
        }
    }

    return 'Misc'
}

function Get-DestinationFolderName {
    param(
        [Parameter(Mandatory)]
        [string]$FileName,

        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Extension,

        [Parameter(Mandatory)]
        [ValidateSet('Extension', 'Category')]
        [string]$GroupingMode
    )

    if ($GroupingMode -eq 'Category') {
        return Get-CategoryFolderName -FileName $FileName -Extension $Extension
    }

    return Get-ExtensionFolderName -Extension $Extension
}

function Get-UniqueDestinationPath {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return $Path
    }

    $directory = Split-Path -Path $Path -Parent
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    $extension = [System.IO.Path]::GetExtension($Path)
    $counter = 1

    while ($true) {
        $candidate = Join-Path -Path $directory -ChildPath ("{0} ({1}){2}" -f $stem, $counter, $extension)
        if (-not (Test-Path -LiteralPath $candidate)) {
            return $candidate
        }

        $counter++
    }
}

function Remove-DirectoryToRecycleBin {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory(
        $Path,
        [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs,
        [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin,
        [Microsoft.VisualBasic.FileIO.UICancelOption]::ThrowException
    )
}

function Test-IsDescendantPath {
    param(
        [Parameter(Mandatory)]
        [string]$RootPath,

        [Parameter(Mandatory)]
        [string]$CandidatePath
    )

    $normalizedRoot = [System.IO.Path]::GetFullPath($RootPath).TrimEnd('\')
    $normalizedCandidate = [System.IO.Path]::GetFullPath($CandidatePath).TrimEnd('\')
    return $normalizedCandidate.StartsWith(
        $normalizedRoot + [System.IO.Path]::DirectorySeparatorChar,
        [System.StringComparison]::OrdinalIgnoreCase
    )
}

function Test-IsFileSystemRoot {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $normalizedPath = [System.IO.Path]::GetFullPath($Path).TrimEnd('\')
    $normalizedRoot = [System.IO.Path]::GetPathRoot($normalizedPath).TrimEnd('\')
    return [System.String]::Equals(
        $normalizedPath,
        $normalizedRoot,
        [System.StringComparison]::OrdinalIgnoreCase
    )
}

function Test-IsSystemDriveRoot {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-IsFileSystemRoot -Path $Path)) {
        return $false
    }

    $systemDrive = [System.Environment]::GetEnvironmentVariable('SystemDrive')
    if ([string]::IsNullOrWhiteSpace($systemDrive)) {
        return $false
    }

    $normalizedPath = [System.IO.Path]::GetFullPath($Path).TrimEnd('\')
    $normalizedSystemDrive = [System.IO.Path]::GetPathRoot($systemDrive + '\').TrimEnd('\')
    return [System.String]::Equals(
        $normalizedPath,
        $normalizedSystemDrive,
        [System.StringComparison]::OrdinalIgnoreCase
    )
}

function Get-FilesAndDirectoriesRecursively {
    param(
        [Parameter(Mandatory)]
        [System.IO.DirectoryInfo]$Directory,

        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[System.IO.FileInfo]]$Files,

        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[System.IO.DirectoryInfo]]$Directories,

        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.HashSet[string]]$SkippedDirectories,

        [switch]$IsRoot
    )

    try {
        $items = Get-ChildItem -LiteralPath $Directory.FullName -Force -ErrorAction Stop
    }
    catch {
        if ($IsRoot) {
            throw
        }

        $null = $SkippedDirectories.Add($Directory.FullName)
        return
    }

    foreach ($item in $items) {
        if ($item.PSIsContainer) {
            if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                continue
            }

            $Directories.Add($item)
            Get-FilesAndDirectoriesRecursively -Directory $item -Files $Files -Directories $Directories -SkippedDirectories $SkippedDirectories
            continue
        }

        $Files.Add($item)
    }
}

function Move-FolderFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$FolderPath,

        [Parameter()]
        [ValidateSet('Extension', 'Category')]
        [string]$GroupingMode = 'Extension'
    )

    if (-not (Test-Path -LiteralPath $FolderPath)) {
        throw "Folder does not exist: $FolderPath"
    }

    $resolvedFolder = (Resolve-Path -LiteralPath $FolderPath).Path
    $folderItem = Get-Item -LiteralPath $resolvedFolder

    if (-not $folderItem.PSIsContainer) {
        throw "Path is not a folder: $resolvedFolder"
    }

    if (Test-IsSystemDriveRoot -Path $resolvedFolder) {
        throw "Sorting the Windows system drive root ($resolvedFolder) is blocked. Choose another folder or a non-system drive."
    }

    $files = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
    $directories = [System.Collections.Generic.List[System.IO.DirectoryInfo]]::new()
    $skippedDirectories = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    $skippedFiles = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    $extensionFoldersUsed = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    $createdFolders = 0
    $movedFiles = 0
    $deletedFolders = 0

    Get-FilesAndDirectoriesRecursively -Directory $folderItem -Files $files -Directories $directories -SkippedDirectories $skippedDirectories -IsRoot

    foreach ($file in $files) {
        try {
            $extensionFolder = Get-DestinationFolderName -FileName $file.Name -Extension $file.Extension -GroupingMode $GroupingMode
            $destinationFolder = Join-Path -Path $resolvedFolder -ChildPath $extensionFolder

            if (-not (Test-Path -LiteralPath $destinationFolder)) {
                $null = New-Item -Path $destinationFolder -ItemType Directory -ErrorAction Stop
                $createdFolders++
            }

            $null = $extensionFoldersUsed.Add($destinationFolder)

            $destinationPath = Join-Path -Path $destinationFolder -ChildPath $file.Name
            if ([System.String]::Equals($file.FullName, $destinationPath, [System.StringComparison]::OrdinalIgnoreCase)) {
                continue
            }

            $uniqueDestination = Get-UniqueDestinationPath -Path $destinationPath
            Move-Item -LiteralPath $file.FullName -Destination $uniqueDestination -ErrorAction Stop
            $movedFiles++
        }
        catch {
            $null = $skippedFiles.Add($file.FullName)
        }
    }

    foreach ($directory in ($directories | Sort-Object { $_.FullName.Length } -Descending)) {
        if ($skippedDirectories.Contains($directory.FullName)) {
            continue
        }

        if ($extensionFoldersUsed.Contains($directory.FullName)) {
            continue
        }

        if (-not (Test-IsDescendantPath -RootPath $resolvedFolder -CandidatePath $directory.FullName)) {
            throw "Refusing to delete a folder outside the selected root: $($directory.FullName)"
        }

        try {
            if (Test-Path -LiteralPath $directory.FullName) {
                Remove-DirectoryToRecycleBin -Path $directory.FullName
                $deletedFolders++
            }
        }
        catch {
            $null = $skippedDirectories.Add($directory.FullName)
        }
    }

    [pscustomobject]@{
        MovedFiles = $movedFiles
        CreatedFolders = $createdFolders
        DeletedFolders = $deletedFolders
        RemainingExtensionFolders = $extensionFoldersUsed.Count
        SkippedDirectories = $skippedDirectories.Count
        SkippedFiles = $skippedFiles.Count
        GroupingMode = $GroupingMode
    }
}

function Move-FolderFilesByExtension {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$FolderPath
    )

    Move-FolderFiles -FolderPath $FolderPath -GroupingMode 'Extension'
}

function Move-FolderFilesByCategory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$FolderPath
    )

    Move-FolderFiles -FolderPath $FolderPath -GroupingMode 'Category'
}

Export-ModuleMember -Function Get-ExtensionFolderName, Get-CategoryFolderName, Get-UniqueDestinationPath, Move-FolderFiles, Move-FolderFilesByExtension, Move-FolderFilesByCategory
