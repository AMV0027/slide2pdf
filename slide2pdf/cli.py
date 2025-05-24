import os
import comtypes.client
import argparse

def convert_all_ppt_in_folder(folder_path):
    folder_path = os.path.abspath(folder_path)
    output_folder = os.path.join(folder_path, "p2pdf")
    os.makedirs(output_folder, exist_ok=True)

    ppt_files = [f for f in os.listdir(folder_path) if f.lower().endswith((".ppt", ".pptx"))]
    if not ppt_files:
        print("❌ No PPT or PPTX files found.")
        return

    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1

    for file in ppt_files:
        input_file = os.path.join(folder_path, file)
        output_file = os.path.join(output_folder, os.path.splitext(file)[0] + ".pdf")
        try:
            presentation = powerpoint.Presentations.Open(input_file, WithWindow=False)
            presentation.SaveAs(output_file, 32)
            presentation.Close()
            print(f"✅ Converted: {file}")
        except Exception as e:
            print(f"❌ Error converting {file}: {e}")

    powerpoint.Quit()
    print(f"\n🎉 Done! PDFs saved in: {output_folder}")

def main():
    parser = argparse.ArgumentParser(description="Convert all PPT/PPTX files in a folder to PDF.")
    parser.add_argument('--path', type=str, help='Path to the folder containing PPT files')
    args = parser.parse_args()

    folder = args.path if args.path else os.getcwd()
    convert_all_ppt_in_folder(folder)
