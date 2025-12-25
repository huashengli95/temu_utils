import os
from pypdf import PdfReader, PdfWriter

def merge_pdfs(folder_path, output_file="merged.pdf"):
    writer = PdfWriter()

    pdf_files = sorted(
        f for f in os.listdir(folder_path)
        if f.lower().endswith(".pdf")
    )

    if not pdf_files:
        print("❌ 文件夹里没有 PDF 文件")
        return

    for pdf in pdf_files:
        pdf_path = os.path.join(folder_path, pdf)
        reader = PdfReader(pdf_path)

        for page in reader.pages:
            writer.add_page(page)

        print(f"✔ 已添加: {pdf}")

    output_path = os.path.join(folder_path, output_file)
    with open(output_path, "wb") as f:
        writer.write(f)

    print(f"\n🎉 合并完成: {output_path}")


if __name__ == "__main__":
    merge_pdfs(os.getcwd())
