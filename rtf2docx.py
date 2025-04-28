import os
import argparse
import pypandoc

def convertRtf2docx(rtfPath, docxPath):
    try:
        # Convert RTF to DOCX
        pypandoc.convert_file(rtfPath, 'docx', outputfile=docxPath)
        print(f"Converted: {rtfPath} -> {docxPath}")
    except Exception as e:
        print(f"Failed to convert {rtfPath}: {e}")

def convertFolder(inputFolder, outputFolder):
    # Create output folder if it doesn't exist
    os.makedirs(outputFolder, exist_ok=True)

    for filename in os.listdir(inputFolder):
        if filename.lower().endswith('.rtf'):
            rtfPath = os.path.join(inputFolder, filename)
            docxFilename = os.path.splitext(filename)[0] + '.docx'
            docxPath = os.path.join(outputFolder, docxFilename)
            convertRtf2docx(rtfPath, docxPath)

def main():
    parser = argparse.ArgumentParser(description='Batch convert RTF files to DOCX using pypandoc.')
    parser.add_argument('inputFolder', type=str, help='Path to the folder containing RTF files.')
    parser.add_argument('outputFolder', type=str, help='Path to the folder to save DOCX files.')

    args = parser.parse_args()

    convertFolder(args.inputFolder, args.outputFolder)
    print("All RTF files have been converted.")

if __name__ == "__main__":
    main()
