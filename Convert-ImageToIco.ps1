function Convert-ImageToIco {
	<#
    .SYNOPSIS
        Converts various image formats to ICO files for use as icons.
    .DESCRIPTION
        The Convert-ImageToIco function converts images from various formats (PNG, JPG, GIF, BMP, TIFF) to ICO format.
        It supports both local files and URLs as input, and allows you to specify the size of the output icon.
        The function uses high-quality scaling to ensure the output icon looks good at the specified size.
        ICO files are commonly used for application icons, favicons for websites, and folder icons in Windows.
        WebP format is not natively supported by the System.Drawing library in .NET. If you need to convert
        WebP images, please convert them to PNG first using an online converter or tools like ImageMagick.
    .PARAMETER Path
        Specifies the path to the source image file or URL.
        Accepts local file paths or URLs starting with http:// or https://.
    .PARAMETER OutputPath
        Specifies the path where the ICO file will be saved.
        If not specified, the function will use the same location and filename as the input file, but with an .ico extension.
    .PARAMETER Size
        Specifies the size of the output icon in pixels (both width and height).
        Acceptable values are 16, 32, 48, 64, 128, 256, or 512.
        The default value is 256.
        Note: For sizes larger than 255 (i.e., 256 and 512), the ICO format uses special handling.
        Some older applications might display these as 256x256 icons.
    .EXAMPLE
        Convert-ImageToIco -Path "C:\Images\logo.png" -Size 32
		Converts the PNG image to a 32x32 ICO file in the same location (C:\Images\logo.ico).
    .EXAMPLE
        Convert-ImageToIco -Path "https://upload.wikimedia.org/wikipedia/commons/e/e0/Check_green_icon.png" -Size 64 -OutputPath "C:\Icons\checkmark.ico"
        Downloads the PNG image from the URL and converts it to a 64x64 ICO file saved at the specified path.
    .EXAMPLE
        Convert-ImageToIco -Path "C:\Images\photo.jpg" -Size 256
        Converts the JPEG image to a 256x256 ICO file in the same location (C:\Images\photo.ico).
    .EXAMPLE
        Get-ChildItem -Path "C:\Images" -Filter "*.png" | ForEach-Object { Convert-ImageToIco -Path $_.FullName -Size 48 }
        Batch converts all PNG files in the C:\Images directory to 48x48 ICO files.
    .NOTES
        Author: Harze2k
        Version: 1.0 (Initial release.)
        Date: 2025-05-21

        Supported image formats: PNG, JPG/JPEG, BMP, GIF, TIFF/TIF
        WebP format is NOT supported natively - convert WebP to PNG first.

        Requirements:
        - Windows PowerShell 5.1 or PowerShell Core 6.0+
        - .NET Framework or .NET Core (for System.Drawing)
    .LINK
        https://learn.microsoft.com/en-us/dotnet/api/system.drawing.image

    .INPUTS
        System.String

    .OUTPUTS
        System.String
        Returns the full path to the created ICO file.
    #>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0)][string]$Path,
		[Parameter(Position = 1)][string]$OutputPath = $null,
		[Parameter()][ValidateSet(16, 32, 48, 64, 128, 256, 512)][int]$Size = 256
	)
	begin {
		# Load necessary assemblies
		Add-Type -AssemblyName System.Drawing
		Add-Type -AssemblyName System.IO
		# Define supported file extensions
		$supportedExtensions = @('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff', '.tif')
		$webpExtension = '.webp'
		# Helper function to determine if the path is a URL
		function Test-IsUrl {
			[CmdletBinding()]
			param (
				[string]$Url
			)
			return $Url -match '^https?://'
		}
		# Helper function to validate file extension
		function Test-SupportedImageFormat {
			[CmdletBinding()]
			param (
				[string]$FilePath
			)
			$extension = [System.IO.Path]::GetExtension($FilePath).ToLower()
			if ($supportedExtensions -contains $extension) {
				return $true
			}
			elseif ($extension -eq $webpExtension) {
				Write-Warning "WebP format is not natively supported by System.Drawing. Consider converting to PNG first."
				return $false
			}
			return $false
		}
	}
	process {
		try {
			Write-Host "Converting image to ICO with size $Size x $Size..."
			# Create a bitmap from either URL or local file
			$originalImage = $null
			if (Test-IsUrl -Url $Path) {
				Write-Host "Downloading image from URL: $Path"
				# Validate file extension for URL
				if (-not (Test-SupportedImageFormat -FilePath $Path)) {
					# If it's WebP, provide specific error
					if ([System.IO.Path]::GetExtension($Path).ToLower() -eq '.webp') {
						Write-Error "WebP format is not natively supported by System.Drawing. Please convert the WebP file to PNG first using an online converter or a tool like ImageMagick."
					}
					Write-Warning "URL may not point to a supported image format. Supported formats: $($supportedExtensions -join ', ')"
					Write-Warning "Attempting to download and convert anyway, but this might fail."
				}
				# Download the image with improved error handling
				try {
					$webClient = New-Object System.Net.WebClient
					# Add a user agent to avoid 403 errors
					$webClient.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
					$imageBytes = $webClient.DownloadData($Path)
					$memStream = New-Object System.IO.MemoryStream($imageBytes, 0, $imageBytes.Length)
					# Try to load as image
					try {
						$originalImage = [System.Drawing.Image]::FromStream($memStream)
					}
					catch {
						Write-Error "Failed to load image from URL. Ensure the URL points to a valid image file. Error: $_"
						return
					}
				}
				catch {
					Write-Error "Failed to download image from URL: $_"
					return
				}
			}
			else {
				# Resolve to absolute path if needed
				$resolvedPath = Resolve-Path $Path -ErrorAction Stop
				Write-Host "Loading image from local path: $resolvedPath"
				# Validate file extension for local file
				if (-not (Test-SupportedImageFormat -FilePath $resolvedPath)) {
					Write-Warning "File extension not recognized as a supported image format. Supported formats: $($supportedExtensions -join ', ')"
					Write-Warning "Attempting to load and convert anyway, but this might fail."
				}
				# Try to load the image
				try {
					$originalImage = [System.Drawing.Image]::FromFile($resolvedPath)
				}
				catch {
					Write-Error "Failed to load image from file. Ensure the file is a valid image. Error: $_"
					return
				}
			}
			# Determine output path if not specified
			if (-not $OutputPath) {
				if (Test-IsUrl -Url $Path) {
					$fileName = $Path.Split('/')[-1].Split('?')[0]
					$OutputPath = Join-Path -Path (Get-Location) -ChildPath ([System.IO.Path]::GetFileNameWithoutExtension($fileName) + '.ico')
				}
				else {
					$OutputPath = [System.IO.Path]::ChangeExtension($Path, '.ico')
				}
			}
			# Create a new bitmap with the desired size
			$bitmap = New-Object System.Drawing.Bitmap($Size, $Size)
			$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
			# Set maximum quality scaling settings
			$graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
			$graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
			$graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
			$graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
			# Draw the original image into the bitmap
			$graphics.DrawImage($originalImage, 0, 0, $Size, $Size)
			# Create ICO file
			# Create a memory stream to hold the ICO file
			$ms = New-Object System.IO.MemoryStream
			$bw = New-Object System.IO.BinaryWriter($ms)
			# Get image bytes
			$imgMs = New-Object System.IO.MemoryStream
			$bitmap.Save($imgMs, [System.Drawing.Imaging.ImageFormat]::Png)
			$imgBytes = $imgMs.ToArray()
			$imgMs.Close()
			# Calculate offsets
			$headerSize = 6 + 16  # Header (6 bytes) + 1 directory entry (16 bytes)
			# Write ICO header
			$bw.Write([UInt16]0) # Reserved, must be 0
			$bw.Write([UInt16]1) # Image type: 1 for ICO
			$bw.Write([UInt16]1) # Number of images (just 1)
			# Write directory entry
			# For sizes > 255, we should use 0 in the header but rely on PNG data for actual dimensions
			$sizeValue = if ($Size -gt 255) { 0 } else { $Size } # 0 means 256+ in ICO format
			$bw.Write([Byte]$sizeValue) # Width (0 means 256 in traditional ICO, but apps will read actual size from PNG)
			$bw.Write([Byte]$sizeValue) # Height (0 means 256 in traditional ICO, but apps will read actual size from PNG)
			$bw.Write([Byte]0) # Color palette
			$bw.Write([Byte]0) # Reserved
			$bw.Write([UInt16]1) # Color planes
			$bw.Write([UInt16]32) # Bits per pixel
			$bw.Write([UInt32]$imgBytes.Length) # Size of image data
			$bw.Write([UInt32]$headerSize) # Offset to image data
			# Write image data (PNG format preserves the actual dimensions)
			$bw.Write($imgBytes)
			# Save the ICO file
			$bytes = $ms.ToArray()
			[System.IO.File]::WriteAllBytes($OutputPath, $bytes)
			# Close streams
			$bw.Close()
			$ms.Close()
			# Cleanup
			$graphics.Dispose()
			$bitmap.Dispose()
			$originalImage.Dispose()
			Write-Host "ICO file created successfully at: $OutputPath"
			return $OutputPath
		}
		catch {
			Write-Error "Error converting image to ICO: $_"
		}
	}
}
# Example usage:
#Convert-ImageToIco -Path "https://i.redd.it/ibtwom3ku12f1.png" -OutputPath "C:\Users\Martin\Desktop\zen_icon.ico" -Size 512