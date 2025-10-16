import fitz  # PyMuPDF
import os
from PIL import Image

def extract_images_from_pdf(pdf_path, output_folder, export_pages=False, dpi=200):
    """
    Extracts embedded images from PDF pages and optionally exports each page as a PNG.
    
    Args:
        pdf_path (str): Path to the PDF file.
        output_folder (str): Folder where extracted images will be saved.
        export_pages (bool): If True, also exports each page as PNG.
        dpi (int): Resolution for page exports (if enabled).
    """
    os.makedirs(output_folder, exist_ok=True)
    doc = fitz.open(pdf_path)
    print(f"Opened PDF: {pdf_path} with {len(doc)} pages.")

    total_images = 0
    for i in range(len(doc)):
        page = doc[i]
        page_index = i + 1
        page_prefix = f"page-{page_index:03d}"

        # --- Extract embedded images ---
        images = page.get_images(full=True)
        if images:
            for img_index, img in enumerate(images, start=1):
                xref = img[0]
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                image_ext = base_image["ext"]
                image_filename = f"{page_prefix}-img-{img_index}.{image_ext}"
                image_path = os.path.join(output_folder, image_filename)
                with open(image_path, "wb") as f:
                    f.write(image_bytes)
                total_images += 1
            print(f"✅ Extracted {len(images)} image(s) from {page_prefix}")
        else:
            print(f"⚠️ No embedded images found on {page_prefix}")

        # --- Optional: export full page as PNG ---
        if export_pages:
            pix = page.get_pixmap(dpi=dpi)
            page_png = os.path.join(output_folder, f"{page_prefix}.png")
            pix.save(page_png)
            print(f"🖼️ Exported page as PNG → {page_png}")

    doc.close()
    print(f"\n🎉 Done! Extracted {total_images} images in total.")
    print(f"All files saved in: {os.path.abspath(output_folder)}")


if __name__ == "__main__":
    # Example usage
    pdf_file = "Recommendation_Letter_Hamza.pdf"           # Path to your PDF file
    output_dir = "test_output"       # Output directory
    extract_images_from_pdf(pdf_file, output_dir, export_pages=True, dpi=200)