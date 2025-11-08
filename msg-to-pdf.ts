#!/usr/bin/env bun
import { resolve, join, basename, dirname } from "node:path";
import { mkdir, readdir, readFile, writeFile, stat } from "node:fs/promises";
import { parse as parseMsgFile } from "@molotochok/msg-viewer/msg-parser";
import puppeteer from "puppeteer";
import { PDFDocument } from "pdf-lib";

interface ConversionResult {
  html: string;
  pdfAttachments: Array<{ fileName: string; content: Uint8Array }>;
  imageAttachments: Array<{ fileName: string; att: any }>;
}

async function convertMessageToHtml(msgData: any): Promise<ConversionResult> {
  let bodyHtml = "";

  // @molotochok/msg-viewer returns content in msgData.content
  const content = msgData.content || msgData;
  
  if (content.bodyHTML) {
    bodyHtml = content.bodyHTML;
  } else if (content.body) {
    // Don't wrap in <pre> if we're going to replace [image] placeholders with actual images
    bodyHtml = content.body.replace(/\n/g, '<br>');
  }

  const attachmentMap = new Map<string, any>();
  const attachmentByFileName = new Map<string, any>();
  const inlineContentIds = new Set<string>();
  const pdfAttachments: Array<{ fileName: string; content: Uint8Array }> = [];

  if (msgData.attachments) {
    for (const att of msgData.attachments) {
      const fileName = att.fileName || att.name || "";
      const attContent = att.content;

      if (fileName.toLowerCase().endsWith(".pdf") && attContent) {
        // Convert DataView to Uint8Array
        const uint8Array = new Uint8Array(attContent.buffer, attContent.byteOffset, attContent.byteLength);
        pdfAttachments.push({
          fileName,
          content: uint8Array,
        });
      }

      if (att.contentId) {
        const cleanContentId = att.contentId.replace(/^<|>$/g, "");
        attachmentMap.set(cleanContentId, att);
      }

      if (fileName) {
        attachmentByFileName.set(fileName, att);
      }
    }
  }

  // Replace cid: references
  bodyHtml = bodyHtml.replace(
    /<img([^>]+)src="cid:([^"]+)"([^>]*)>/gi,
    (match, before, cid, after) => {
      const att = attachmentMap.get(cid);
      if (att && att.content) {
        inlineContentIds.add(cid);
        const data = Buffer.from(att.content).toString("base64");
        const mimeType = att.mimeType || att.contentType || "image/png";
        const dataUri = `data:${mimeType};base64,${data}`;
        return `<img${before}src="${dataUri}"${after}>`;
      }
      return match;
    }
  );

  // Collect image placeholders and their attachments for bottom section
  const imageAttachments: Array<{ fileName: string; att: any }> = [];
  
  // Remove [imageX.ext] placeholders from body and collect them
  bodyHtml = bodyHtml.replace(
    /\[([^\]]+\.(png|jpg|jpeg|gif|bmp|tif|tiff))\]/gi,
    (match, fileName) => {
      let att = attachmentByFileName.get(fileName);

      if (!att) {
        const lowerFileName = fileName.toLowerCase();
        for (const [key, value] of attachmentByFileName.entries()) {
          if (key.toLowerCase() === lowerFileName) {
            att = value;
            break;
          }
        }
      }

      if (att && att.content) {
        inlineContentIds.add(att.name || att.fileName || fileName);
        imageAttachments.push({ fileName, att });
        return ""; // Remove placeholder from body
      }
      return match;
    }
  );

  const headers = `
    <div style="font-family: Arial, sans-serif; margin-bottom: 20px; border-bottom: 2px solid #ccc; padding-bottom: 10px;">
      <h2 style="margin: 0 0 10px 0;">${content.subject || "(No Subject)"}</h2>
      <p style="margin: 5px 0;"><strong>From:</strong> ${content.senderName || content.senderEmail || ""}</p>
      <p style="margin: 5px 0;"><strong>To:</strong> ${msgData.recipients?.map((r: any) => r.displayName || r.email).join(", ") || ""}</p>
      ${content.cc ? `<p style="margin: 5px 0;"><strong>Cc:</strong> ${content.cc}</p>` : ""}
      <p style="margin: 5px 0;"><strong>Date:</strong> ${content.date || ""}</p>
    </div>
  `;

  // Build attachments section at the bottom with ALL images and embedded PDFs
  let attachmentsList = "";
  
  // Collect ALL image attachments (both referenced and unreferenced)
  const allImageAttachments: Array<{ fileName: string; att: any }> = [];
  if (msgData.attachments) {
    for (const att of msgData.attachments) {
      const fileName = att.fileName || att.name || "";
      const isImage = /\.(png|jpg|jpeg|gif|bmp|tif|tiff)$/i.test(fileName);
      
      if (isImage && att.content) {
        // Check if already in imageAttachments (from body placeholders)
        const alreadyAdded = imageAttachments.some(img => img.fileName === fileName);
        if (!alreadyAdded) {
          allImageAttachments.push({ fileName, att });
        }
      }
    }
  }
  
  // Combine placeholder images with unreferenced images
  const allImages = [...imageAttachments, ...allImageAttachments];
  
  // Show all images if any
  if (allImages.length > 0) {
    const imagesHtml = allImages.map(({ fileName, att }) => {
      const uint8Array = new Uint8Array(att.content.buffer, att.content.byteOffset, att.content.byteLength);
      const data = Buffer.from(uint8Array).toString("base64");
      const mimeType = att.mimeType || `image/${fileName.split(".").pop()}`;
      const dataUri = `data:${mimeType};base64,${data}`;
      return `
        <div style="margin: 20px 0;">
          <img src="${dataUri}" style="max-width: 100%; height: auto; display: block;" />
          <div style="font-weight: bold; margin-top: 5px; text-align: center; color: #666;">[${fileName}]</div>
        </div>
      `;
    }).join("\n");
    
    attachmentsList = `
      <div style="font-family: Arial, sans-serif; margin-top: 40px; padding-top: 20px;">
        <div style="text-align: center; font-size: 20px; color: #666; margin-bottom: 15px;">* * *</div>
        <h3 style="margin: 10px 0;">Attachments:</h3>
        <div>${imagesHtml}</div>
      </div>
    `;
    
    // Mark all shown images as inline so they don't appear in files list
    allImages.forEach(({ fileName, att }) => {
      inlineContentIds.add(att.name || att.fileName || fileName);
    });
  }
  
  // Add non-inline file attachments list if any
  if (msgData.attachments && msgData.attachments.length > 0) {
    const nonInlineAttachments = msgData.attachments.filter((att: any) => {
      const fileName = att.name || att.fileName || "";
      if (inlineContentIds.has(fileName)) return false;
      if (fileName.toLowerCase().endsWith(".pdf")) return false;
      if (att.contentId) {
        const cleanContentId = att.contentId.replace(/^<|>$/g, "");
        if (inlineContentIds.has(cleanContentId)) return false;
      }
      return true;
    });

    if (nonInlineAttachments.length > 0) {
      const filesListHtml = `
        <div style="margin-top: 20px;">
          <h4 style="margin: 10px 0;">Files:</h4>
          <ul style="margin: 10px 0; padding-left: 30px;">
            ${nonInlineAttachments.map((att: any) => `<li style="margin: 5px 0;">${att.name || att.fileName || "Unnamed attachment"}</li>`).join("")}
          </ul>
        </div>
      `;
      
      if (imageAttachments.length === 0) {
        // No images, so create the attachments section with just files
        attachmentsList = `
          <div style="font-family: Arial, sans-serif; margin-top: 40px; padding-top: 20px;">
            <div style="text-align: center; font-size: 20px; color: #666; margin-bottom: 15px;">* * *</div>
            <h3 style="margin: 10px 0;">Attachments:</h3>
            ${filesListHtml}
          </div>
        `;
      } else {
        // Append files list to existing attachments section
        attachmentsList = attachmentsList.replace('</div>\n    ', filesListHtml + '</div>\n    ');
      }
    }
  }

  const fullHtml = `
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="UTF-8">
        <style>
          body { font-family: Arial, sans-serif; padding: 20px; max-width: 800px; margin: 0 auto; }
          img { max-width: 100%; height: auto; }
        </style>
      </head>
      <body>
        ${headers}
        <div class="body-content">
          ${bodyHtml}
        </div>
        ${attachmentsList}
      </body>
    </html>
  `;

  return { html: fullHtml, pdfAttachments, imageAttachments: allImages };
}

async function createPdfFromHtml(
  htmlContent: string,
  outputFilePath: string,
  pdfAttachments: Array<{ fileName: string; content: Uint8Array }> = []
): Promise<void> {
  const browser = await puppeteer.launch({
    headless: true,
    args: ["--no-sandbox", "--disable-setuid-sandbox"],
  });

  const page = await browser.newPage();
  
  await page.setContent(htmlContent, { waitUntil: "networkidle0" });
  
  const pdfBuffer = await page.pdf({ format: "A4", printBackground: true });

  if (pdfAttachments.length > 0) {
    const mergedPdf = await PDFDocument.create();
    const mainPdf = await PDFDocument.load(pdfBuffer);
    const mainPages = await mergedPdf.copyPages(mainPdf, mainPdf.getPageIndices());
    mainPages.forEach((page) => mergedPdf.addPage(page));

    for (const attachment of pdfAttachments) {
      try {
        // Create a label page for the PDF using Puppeteer with proper font for Chinese
        const labelHtml = `
          <!DOCTYPE html>
          <html>
            <head>
              <meta charset="UTF-8">
              <style>
                @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@700&display=swap');
                * { margin: 0; padding: 0; box-sizing: border-box; }
                body { 
                  font-family: 'Noto Sans SC', Arial, sans-serif; 
                  padding: 40px;
                  background: white;
                }
                .label {
                  font-size: 16px;
                  font-weight: bold;
                  padding: 20px;
                  border: 2px solid #333;
                  border-radius: 8px;
                  background: #f5f5f5;
                  word-break: break-word;
                }
              </style>
            </head>
            <body>
              <div class="label">[${attachment.fileName}]</div>
            </body>
          </html>
        `;
        
        await page.setContent(labelHtml, { waitUntil: "networkidle2" });
        // Wait extra time for Google Fonts to load
        await page.evaluate(() => document.fonts.ready);
        await new Promise(resolve => setTimeout(resolve, 1000));
        
        const labelPdfBuffer = await page.pdf({ 
          format: "A4", 
          printBackground: true,
          preferCSSPageSize: false
        });
        
        const labelPdf = await PDFDocument.load(labelPdfBuffer);
        const labelPages = await mergedPdf.copyPages(labelPdf, labelPdf.getPageIndices());
        labelPages.forEach((p) => mergedPdf.addPage(p));
        
        // Then add the PDF pages
        const attachedPdf = await PDFDocument.load(attachment.content);
        const attachedPages = await mergedPdf.copyPages(attachedPdf, attachedPdf.getPageIndices());
        attachedPages.forEach((p) => mergedPdf.addPage(p));
      } catch (error) {
        console.error(`Failed to merge PDF attachment ${attachment.fileName}:`, error instanceof Error ? error.message : String(error));
      }
    }

    const mergedPdfBytes = await mergedPdf.save();
    await writeFile(outputFilePath, mergedPdfBytes);
  } else {
    await writeFile(outputFilePath, pdfBuffer);
  }
  
  await browser.close();
}

async function processFile(msgFilePath: string) {
  const file = basename(msgFilePath);
  const dir = dirname(msgFilePath);
  const pdfFileName = basename(file, ".msg") + ".pdf";
  const pdfFilePath = join(dir, pdfFileName);

  console.log(`${msgFilePath} -> ${pdfFilePath}`);

  const fileBuffer = await readFile(msgFilePath);
  const msgData = parseMsgFile(new DataView(fileBuffer.buffer));

  const result = await convertMessageToHtml(msgData);
  await createPdfFromHtml(result.html, pdfFilePath, result.pdfAttachments);
}

async function main() {
  const argv = process.argv.slice(2);
  
  const showHelp = () => {
    console.log(`
Usage: msg-to-pdf [options]

Options:
  -d, --directory <dir>   Convert all .msg files in directory
  -f, --file <path>       Convert a single .msg file
  -h, --help              Show this help message

Examples:
  msg-to-pdf -d ./msgs              # Convert all .msg files in ./msgs/
  msg-to-pdf -f "email.msg"         # Convert single file to same directory
    `);
    process.exit(0);
  };

  // Simple argument parser
  let directory: string | undefined;
  let file: string | undefined;
  
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === "-h" || arg === "--help") {
      showHelp();
    } else if (arg === "-d" || arg === "--directory") {
      directory = argv[++i];
    } else if (arg === "-f" || arg === "--file") {
      file = argv[++i];
    }
  }

  // Show error if no arguments provided
  if (!file && !directory) {
    console.error("Error: Please provide either -d <directory> or -f <file>\n");
    showHelp();
  }

  // Single file mode
  if (file) {
    const msgFilePath = resolve(process.cwd(), file);
    try {
      const stats = await stat(msgFilePath);
      if (!stats.isFile()) {
        console.error(`Error: ${file} is not a file`);
        process.exit(1);
      }
      await processFile(msgFilePath);
    } catch (error) {
      console.error(`Error processing ${file}:`, error instanceof Error ? error.message : String(error));
      process.exit(1);
    }
    return;
  }

  // Directory mode
  const inputDir = resolve(process.cwd(), directory!);
  
  try {
    const files = await readdir(inputDir);
    const msgFiles = files.filter(f => f.endsWith(".msg"));

    if (msgFiles.length === 0) {
      console.log(`No .msg files found in ${inputDir}`);
      return;
    }

    for (const file of msgFiles) {
      try {
        const msgFilePath = join(inputDir, file);
        await processFile(msgFilePath);
      } catch (error) {
        console.error(`Error processing ${file}:`, error instanceof Error ? error.message : String(error));
      }
    }
  } catch (error) {
    console.error(`Error reading directory ${inputDir}:`, error instanceof Error ? error.message : String(error));
    process.exit(1);
  }
}

if (import.meta.main) {
  main();
}

