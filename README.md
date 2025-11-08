# Outlook MSG to PDF CLI

Convert Outlook `.msg` files to self-contained PDFs with embedded images and merged PDF attachments.

## Features

- 📧 Email metadata (Subject, From, To, Date)
- 🖼️ All images embedded at bottom with filename labels
- 📎 PDF attachments automatically merged
- 🌏 Unicode support (Chinese characters, etc.)
- ⚡ Fast conversion with Puppeteer
- 📦 Standalone binary (includes Chromium)

## Download

Download pre-built binaries from the [Releases](../../releases) page:
- `msg-to-pdf-linux-x64` - Linux x64
- `msg-to-pdf-macos-arm64` - macOS Apple Silicon
- `msg-to-pdf-macos-x64` - macOS Intel
- `msg-to-pdf-windows-x64.exe` - Windows x64

Or build from source:

```bash
bun install
bun run build
```

The compiled binary will be in `build/msg-to-pdf`.

## Usage

```bash
# Convert all .msg files in a directory
./build/msg-to-pdf -d ./msgs

# Convert a single file
./build/msg-to-pdf -f "email.msg"

# Show help
./build/msg-to-pdf --help
```

PDFs are created in the same folder as the input `.msg` files.

## Development

```bash
# Run without compiling
bun run start

# Rebuild binary
bun run build
```

## Output Format

Each PDF contains:
1. Email header (Subject, From, To, Date)
2. Email body text
3. `* * *` separator
4. **Attachments:** section with:
   - All embedded images with `[filename]` labels below
   - Label pages for PDF attachments (with Unicode support)
   - Merged PDF attachment content

## Example

```bash
./build/msg-to-pdf -f "Container contents.msg"
```

Creates `Container contents.pdf` in the same directory with all images and attachments embedded.

## How to Create a Release

1. Create and push a git tag:
```bash
git tag v1.0.0
git push origin v1.0.0
```

2. GitHub Actions will automatically build binaries for all platforms and create a release

3. Download the binaries from the Releases page

