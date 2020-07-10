from mailmerge import MailMerge
import tkinter as tk
from tkinter import ttk
import time
import os
import sys
import json
import subprocess
from pathlib import Path
from tqdm.auto import tqdm
import PyPDF2
from python_autocite.lib import formatter
import python_autocite.lib

currentdir = os.getcwd()

#python_autocite.lib.formatter.APAFormatter.format("something.docx", "https://github.com/thenaterhood/python-autocite/archive/")

try:
    # 3.8+
    from importlib.metadata import version
except ImportError:
    #rom importlib_metadata import version
    print("import error")

__version__ = version(__package__)

def windows(paths, keep_active):
    import win32com.client
    #print(paths)
    word = win32com.client.Dispatch("Word.Application")
    wdFormatPDF = 17

    if paths["batch"]:
        for docx_filepath in tqdm(sorted(Path(paths["input"]).glob("*.docx"))):
            pdf_filepath = Path(paths["output"]) / (str(docx_filepath.stem) + ".pdf")
            doc = word.Documents.Open(str(docx_filepath))
            doc.SaveAs(str(pdf_filepath), FileFormat=wdFormatPDF)
            doc.Close()
    else:
        pbar = tqdm(total=1)
        docx_filepath = Path(paths["input"]).resolve()
        pdf_filepath = Path(paths["output"]).resolve()
        doc = word.Documents.Open(str(docx_filepath))
        doc.SaveAs(str(pdf_filepath), FileFormat=wdFormatPDF)
        doc.Close()
        pbar.update(1)

    if not keep_active:
        word.Quit()

def windows25(paths, keep_active):
    import win32com.client
    word = win32com.client.Dispatch("Excel.Application")
    wdFormatPDF = 17
    print(paths)
    if paths["batch"]:
        for docx_filepath in tqdm(sorted(Path(paths["input"]).glob("*.xlsx"))):
            pdf_filepath = Path(paths["output"]) / (str(docx_filepath.stem) + ".pdf")
            doc = word.Documents.Open(str(docx_filepath))
            doc.SaveAs(str(pdf_filepath), FileFormat=wdFormatPDF)
            doc.Close()
    else:
        pbar = tqdm(total=1)
        docx_filepath = Path(paths["input"]).resolve()
        pdf_filepath = Path(paths["output"]).resolve()
        doc = word.Documents.Open(str(docx_filepath))
        doc.SaveAs(str(pdf_filepath), FileFormat=wdFormatPDF)
        doc.Close()
        pbar.update(1)

    if not keep_active:
        word.Quit()


def macos(paths, keep_active):
    script = (Path(__file__).parent / "convert.jxa").resolve()
    cmd = [
        "/usr/bin/osascript",
        "-l",
        "JavaScript",
        str(script),
        str(paths["input"]),
        str(paths["output"]),
        str(keep_active).lower(),
    ]

    def run(cmd):
        process = subprocess.Popen(cmd, stderr=subprocess.PIPE)
        while True:
            line = process.stderr.readline().rstrip()
            if not line:
                break
            yield line.decode("utf-8")

    total = len(list(Path(paths["input"]).glob("*.docx"))) if paths["batch"] else 1
    pbar = tqdm(total=total)
    for line in run(cmd):
        try:
            msg = json.loads(line)
        except ValueError:
            continue
        if msg["result"] == "success":
            pbar.update(1)
        elif msg["result"] == "error":
            print(msg)
            sys.exit(1)

def resolve_paths(input_path, output_path):
    input_path = Path(input_path).resolve()
    output_path = Path(output_path).resolve() if output_path else None
    output = {}
    if input_path.is_dir():
        output["batch"] = True
        output["input"] = str(input_path)
        if output_path:
            assert output_path.is_dir()
        else:
            output_path = str(input_path)
        output["output"] = output_path
    else:
        output["batch"] = False
        assert str(input_path).endswith(".docx")
        output["input"] = str(input_path)
        if output_path and output_path.is_dir():
            output_path = str(output_path / (str(input_path.stem) + ".pdf"))
        elif output_path:
            assert str(output_path).endswith(".pdf")
        else:
            output_path = str(input_path.parent / (str(input_path.stem) + ".pdf"))
        output["output"] = output_path
    return output

def convert(input_path, output_path=None, keep_active=False):
    paths = resolve_paths(input_path, output_path)
    if sys.platform == "darwin":
        return macos(paths, keep_active)
    elif sys.platform == "win32":
        return windows(paths, keep_active)
    else:
        raise NotImplementedError(
            "docx2pdf is not implemented for linux as it requires Microsoft Word to be installed"
        )

def cli():

    import textwrap
    import argparse

    if "--version" in sys.argv:
        print(__version__)
        sys.exit(0)

    description = textwrap.dedent(
        """
    Example Usage:

    Convert single docx file in-place from myfile.docx to myfile.pdf:
        docx2pdf myfile.docx

    Batch convert docx folder in-place. Output PDFs will go in the same folder:
        docx2pdf myfolder/

    Convert single docx file with explicit output filepath:
        docx2pdf input.docx output.docx

    Convert single docx file and output to a different explicit folder:
        docx2pdf input.docx output_dir/

    Batch convert docx folder. Output PDFs will go to a different explicit folder:
        docx2pdf input_dir/ output_dir/
    """
    )

    formatter_class = lambda prog: argparse.RawDescriptionHelpFormatter(
        prog, max_help_position=32
    )
    parser = argparse.ArgumentParser(
        description=description, formatter_class=formatter_class
    )
    parser.add_argument(
        "input",
        help="input file or folder. batch converts entire folder or convert single file",
    )
    parser.add_argument("output", nargs="?", help="output file or folder")
    parser.add_argument(
        "--keep-active",
        action="store_true",
        default=False,
        help="prevent closing word after conversion",
    )
    parser.add_argument(
        "--version", action="store_true", default=False, help="display version and exit"
    )

    if len(sys.argv) == 1:
        parser.print_help()
        sys.exit(0)
    else:
        args = parser.parse_args()

    convert(args.input, args.output, args.keep_active)

def popupmsg(msg):
    NORM_FONT = ("Helvetica", 10)
    popup = tk.Tk()
    popup.wm_title("!")
    label = ttk.Label(popup, text=msg, font=NORM_FONT)
    label.pack(side="top", fill="x", pady=10)
    B1 = ttk.Button(popup, text="Okay", command = exit)
    B1.pack()
    popup.mainloop()

def Convertprocess():
    test = tk.Tk()
    test.wm_title("Andrew's Paper Machine")
    tk.Label(test, text="Project Name").grid(row=0)
    projectname = tk.Entry(test)
    projectname.grid(row=0, column=1)
    #project_name = projectname.get()
    def convert2():
        project_name = projectname.get()
        print(project_name)
        convert(f"{currentdir}\\{project_name}.docx", f"{currentdir}\\{project_name}.pdf")
        popupmsg("Document Converted!!")
        time.sleep(4)
        exit()

    tk.Button(test,
              text='Convert',
              command=convert2).grid(row=4,
                                        column=0,
                                        sticky=tk.W,
                                        pady=4)

def Writemailmerge():
    def WriteAssignment(project_name, teacher, idofcourse, idofassignment):
        fullname = "Andrew Maddox"
        schoolname = "University of advancing technology"
        fullname = fullname.title()
        schoolname = schoolname.title()
        print("-" * 45)
        print("Creator - Andrew Maddox")
        print("This is my project management script")
        print("-" * 45)

        template = "Template.docx"

        document = MailMerge(template)
        print(document.get_merge_fields())

        document.merge(
            AssignmentID=f'{idofassignment}',
            IndividualName=f'{fullname}',
            AssignmentName=f'{project_name}',
            Teacher=f'{teacher}',
            ClassID=f'{idofcourse}',)

        document.write(f'{project_name}.docx')
        popupmsg("Document Created!!")

    def show_entry_fields():
        project_name = projectname.get()
        teacher = teacherentry.get()
        idofcourse = idofcourseentry.get()
        idofassignment = idofassignmententry.get()
        print(project_name)
        WriteAssignment(project_name, teacher, idofcourse, idofassignment)
        master.quit()

    master = tk.Tk()
    master.wm_title("Andrew's Paper Machine")

    tk.Label(master, text="Project Name").grid(row=0)
    tk.Label(master, text="Teacher").grid(row=1)
    tk.Label(master, text="Course ID").grid(row=2)
    tk.Label(master, text="Assignment ID").grid(row=3)

    projectname = tk.Entry(master)
    teacherentry = tk.Entry(master)
    idofcourseentry = tk.Entry(master)
    idofassignmententry = tk.Entry(master)

    projectname.grid(row=0, column=1)
    teacherentry.grid(row=1, column=1)
    idofcourseentry.grid(row=2, column=1)
    idofassignmententry.grid(row=3, column=1)

    tk.Button(master,
              text='Quit',
              command=master.quit).grid(row=4,
                                        column=0,
                                        sticky=tk.W,
                                        pady=4)
    tk.Button(master,
              text='Finished', command=show_entry_fields).grid(row=4,
                                                           column=1,
                                                           sticky=tk.W,
                                                           pady=4)

pdfs = []

def PDFMerge():
    global pdfs
    def Destroy2():
        pdfmerge()
    def Destroy():
        pdfname = projectname2.get()
        print(pdfname)
        if ".pdf" not in pdfname:
            pdfname = f"{pdfname}.pdf"
        pdfs.append(pdfname)
        master2.destroy()
        PDFMerge()
    #print("test")
    master2 = tk.Tk()
    master2.wm_title("Andrew's PDF Machine")
    tk.Label(master2, text="Input PDF name").grid(row=0)
    projectname2 = tk.Entry(master2)
    projectname2.grid(row=0, column=1)
    #pdfname = projectname2.get()
    #print(pdfname)
    #if ".pdf" not in pdfname:
    #    pdfname = f"{pdfname}.pdf"
    def pdfmerge():
        def lastmerger():
            global pdfs
            merger = PyPDF2.PdfFileMerger()
            for pdf in pdfs:
                merger.append(pdf)
            pdfmergename = projectname3.get()
            if ".pdf" not in pdfmergename:
                pdfmergename = f"{pdfmergename}.pdf"
            print(pdfmergename)
            merger.write(f"{pdfmergename}")
            merger.close()
            popupmsg("PDF Files Merged!!!")

        master = tk.Tk()
        master.wm_title("Andrew's PDF Machine")
        tk.Label(master, text="Output PDF name").grid(row=0)
        projectname3 = tk.Entry(master)
        projectname3.grid(row=0, column=1)
        #pdfmergename = projectname3.get()

        tk.Button(master,
                  text='Finished', command=lastmerger).grid(row=4,
                                                         column=1,
                                                         sticky=tk.W,
                                                         pady=4)

    #pdfs.append(pdfname)
    #if pdfname == "done":
    #    pdfmerge()
    tk.Button(master2,
              text='Add PDF to list', command=Destroy).grid(row=4,
                                                               column=1,
                                                               sticky=tk.W,
                                                               pady=4)
    tk.Button(master2,
              text='Finished', command=Destroy2).grid(row=4,
                                                     column=2,
                                                     sticky=tk.W,
                                                     pady=4)
    master2.mainloop()

parent = tk.Tk()
frame = tk.Frame(parent)
frame.pack()

text_disp= tk.Button(frame,
                   text="Create Document",
                   command=Writemailmerge
                   )

text_disp.pack(side=tk.LEFT)

exit_button = tk.Button(frame,
                   text="Convert Document to PDF",
                   fg="green",
                   command=Convertprocess)
exit_button.pack(side=tk.RIGHT)
exit_button = tk.Button(frame,
                   text="Merge PDF Files",
                   fg="green",
                   command=PDFMerge)
exit_button.pack(side=tk.RIGHT)


parent.mainloop()
