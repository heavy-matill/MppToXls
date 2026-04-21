#!/usr/bin/env python3
"""
Convert MPP (Microsoft Project) files to XLS using MPXJ library.
Requires: jpype1 and MPXJ JAR files

Usage:
    python mpp_to_xls_converter.py <input_mpp_file> <output_xls_file>
    python mpp_to_xls_converter.py project.mpp project.xls
"""

import sys
import os
from pathlib import Path
try:
    import xlsxwriter
except ImportError:
    xlsxwriter = None


def convert_mpp_to_xls(mpp_file: str, xls_file: str) -> None:
    """
    Convert an MPP file to XLSX format using MPXJ.
    
    Args:
        mpp_file: Path to input MPP file
        xls_file: Path to output XLS/XLSX file
        
    Raises:
        FileNotFoundError: If MPP file doesn't exist or MPXJ JARs not found
        Exception: If conversion fails
    """
    try:
        import jpype
        import jpype.imports
    except ImportError:
        print("Error: jpype1 not installed. Install with: pip install jpype1")
        sys.exit(1)
    
    # Validate input file
    mpp_path = Path(mpp_file)
    if not mpp_path.exists():
        print(f"Error: MPP file not found: {mpp_file}")
        sys.exit(1)
    
    if not mpp_path.suffix.lower() == ".mpp":
        print(f"Warning: File doesn't have .mpp extension: {mpp_file}")
    
    # MPXJ JAR location - check multiple locations
    # 1. First check environment variable
    mpxj_jar_dir = os.environ.get("MPXJ_JAR_PATH")
    if mpxj_jar_dir:
        mpxj_jar_dir = Path(mpxj_jar_dir)
    else:
        # 2. Check relative to script location (development)
        mpxj_jar_dir = Path(__file__).parent / "lib" / "mpxj"
        if not mpxj_jar_dir.exists():
            # 3. Check in current working directory (for packaged exe)
            mpxj_jar_dir = Path.cwd() / "lib" / "mpxj"
        if not mpxj_jar_dir.exists():
            # 4. Check one level up (if exe is in subfolder)
            mpxj_jar_dir = Path.cwd().parent / "lib" / "mpxj"
    
    if not mpxj_jar_dir.exists():
        print(f"Error: MPXJ JAR directory not found at: {mpxj_jar_dir}")
        print("Please download MPXJ from: https://www.taphandle.com/mpxj/")
        print("And place the JAR files in a 'lib/mpxj' directory next to the executable")
        print(f"Or set the MPXJ_JAR_PATH environment variable")
        sys.exit(1)
    
    # Find all JAR files (including dependencies)
    jar_files = list(mpxj_jar_dir.glob("*.jar"))
    jar_files.extend(mpxj_jar_dir.glob("lib/*.jar"))
    if not jar_files:
        print(f"Error: No JAR files found in: {mpxj_jar_dir}")
        sys.exit(1)
    
    try:
        # Start JVM if not already started
        if not jpype.isJVMStarted():
            # Check JAVA_HOME is set
            java_home = os.environ.get("JAVA_HOME")
            if not java_home:
                print("Error: JAVA_HOME environment variable is not set.")
                print("Please install Java JDK and set JAVA_HOME environment variable:")
                print("  Windows: setx JAVA_HOME \"C:\\Program Files\\Java\\jdk-21\"")
                print("  Then restart your terminal/IDE and try again.")
                sys.exit(1)
            
            try:
                # Build classpath string with semicolons for Windows
                classpath = ";".join(str(jar) for jar in jar_files)
                
                # Start JVM with classpath parameter
                jpype.startJVM(classpath=classpath, convertStrings=True)
            except jpype._jvmfinder.JVMNotFoundException as e:
                print(f"Error: Could not find JVM. {e}")
                print(f"JAVA_HOME is set to: {java_home}")
                print("Verify this path contains the Java JDK installation with jvm.dll")
                sys.exit(1)
        
        # Use JClass to load Java classes directly (correct package names)
        File = jpype.JClass("java.io.File")
        MPPReader = jpype.JClass("org.mpxj.mpp.MPPReader")
        
        # Read MPP file
        print(f"Reading MPP file: {mpp_file}")
        reader = MPPReader()
        project = reader.read(File(str(mpp_path.absolute())))
        
        # Get project data
        print("Extracting project data...")
        tasks = list(project.getTasks())
        print(f"Found {len(tasks)} tasks")
        
        # Create Excel file using Python
        # Use xlsx format (Office Open XML) which is more compatible
        xlsx_file = xls_file.replace('.xls', '.xlsx') if not xls_file.endswith('.xlsx') else xls_file
        print(f"Creating Excel file: {xlsx_file}")
        
        if xlsxwriter:
            # Use xlsxwriter if available
            workbook = xlsxwriter.Workbook(xlsx_file)
            worksheet = workbook.add_worksheet("Project")
            
            # Write headers
            headers = ["ID", "Name", "Duration", "Start", "End", "Resources"]
            for col, header in enumerate(headers):
                worksheet.write(0, col, header)
            
            # Write tasks
            for row_idx, task in enumerate(tasks, start=1):
                task_id = task.getID()
                task_name = task.getName()
                duration = task.getDuration()
                start_date = task.getStart()
                end_date = task.getFinish()
                resource_names = task.getResourceNames()
                
                worksheet.write(row_idx, 0, str(task_id) if task_id else "")
                worksheet.write(row_idx, 1, str(task_name) if task_name else "")
                worksheet.write(row_idx, 2, str(duration) if duration else "")
                worksheet.write(row_idx, 3, str(start_date) if start_date else "")
                worksheet.write(row_idx, 4, str(end_date) if end_date else "")
                worksheet.write(row_idx, 5, str(resource_names) if resource_names else "")
            
            workbook.close()
        else:
            # Fallback to CSV if xlsxwriter not available
            import csv
            csv_file = xlsx_file.replace('.xlsx', '.csv')
            print(f"xlsxwriter not installed, saving as CSV instead: {csv_file}")
            with open(csv_file, 'w', newline='') as f:
                writer = csv.writer(f)
                headers = ["ID", "Name", "Duration", "Start", "End", "Resources"]
                writer.writerow(headers)
                for task in tasks:
                    writer.writerow([
                        str(task.getID()) if task.getID() else "",
                        str(task.getName()) if task.getName() else "",
                        str(task.getDuration()) if task.getDuration() else "",
                        str(task.getStart()) if task.getStart() else "",
                        str(task.getFinish()) if task.getFinish() else "",
                        str(task.getResourceNames()) if task.getResourceNames() else ""
                    ])
        
        print("✓ Conversion completed successfully!")
        
    except Exception as e:
        print(f"Error during conversion: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


def main():
    """Main entry point."""
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(1)
    
    mpp_file = sys.argv[1]
    xls_file = sys.argv[2]
    
    convert_mpp_to_xls(mpp_file, xls_file)


if __name__ == "__main__":
    main()
