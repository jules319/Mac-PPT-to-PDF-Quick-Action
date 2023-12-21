import subprocess
import sys
import os

def convert_powerpoint_to_pdf(input_path, output_path):
    # AppleScript code with placeholders for input and output paths
    script = """
    on savePowerPointAsPDF(documentPath, PDFPath)
        set f to POSIX file documentPath as alias
        tell application "Microsoft PowerPoint"
            open f
            set PDFPath to my createEmptyFile(PDFPath)
            delay 1
            save active presentation in PDFPath as save as PDF
            delay 1
            close active presentation
        end tell
    end savePowerPointAsPDF

    on createEmptyFile(f)
        do shell script "touch " & quoted form of POSIX path of f
        return (POSIX path of f) as POSIX file
    end createEmptyFile

    savePowerPointAsPDF("{0}", "{1}")
    """.format(input_path, output_path)
    
    # Run the AppleScript using osascript
    result = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
    if result.returncode != 0:
        print(f"Error: {result.stderr}")
    else:
        print(f"Successfully converted {input_path} to {output_path}")


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print(f"Usage: python {sys.argv[0]} input_file1 [input_file2 ...] /path_to_output_directory/")
        sys.exit(1)
    
    # The last argument is the output directory, all preceding arguments are input files.
    output_directory = sys.argv[-1]
    input_files = sys.argv[1:-1]
    
    # Launch PowerPoint at the beginning
    subprocess.run(["osascript", "-e", 'tell application "Microsoft PowerPoint" to launch'])
    
    for input_file in input_files:
        # Create the output path by joining the output directory with the input filename (without extension) + ".pdf"
        output_file = os.path.join(output_directory, os.path.splitext(os.path.basename(input_file))[0] + ".pdf")
        print(f"Input: {input_file}, Output: {output_file}")
        convert_powerpoint_to_pdf(input_file, output_file)
    
    # Quit PowerPoint at the end
    subprocess.run(["osascript", "-e", 'tell application "Microsoft PowerPoint" to quit'])
