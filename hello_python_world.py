import sys
import time
import asyncio
import pythoncom
import win32com.client

# Solidworks Version (write the year):
SWV = 2023
# API Version
SWAV = SWV - 1992
# Elapsed time
elapsed = 0

async def main():
    # Attempt to connect to SolidWorks
    try:
        sw = win32com.client.GetActiveObject("SldWorks.Application")
    except:
        print("Failed to connect. Launching a new instance.")
        sw = win32com.client.Dispatch("SldWorks.Application")
        sw.Visible = True

    if not sw:
        print("Failed to connect or launch SolidWorks")
        return

    print(f"Solidworks API Version: {SWAV}", "\n", f"Solidworks Version: {SWV}")

    Model = sw.ActiveDoc

    if Model is None:
        print("No active document found. Creating a new document...")
        # Create a new part document (you can change the document type)
        partTemplate = "C:\\ProgramData\\SOLIDWORKS\\SOLIDWORKS " + str(SWV) + "\\templates\\Part.prtdot"
        Model = sw.NewDocument(partTemplate, 0, 0, 0) # You can adjust the last three parameters as needed


    ARG_NULL = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
    ck = Model.Extension.SelectByID2("Front", "PLANE", 0, 0, 0, False, 0, ARG_NULL, 0)
    Model.SketchManager.InsertSketch(True)
    Model.ClearSelection2(True)

    mySketchText = Model.InsertSketchText(
        0, 0, 0, "Hello Python World!", 0, 0, 0, 100, 100
    )
    myFeature = Model.FeatureManager.FeatureExtrusion2(
        True,
        False,
        False,
        0,
        0,
        0.001,
        0.001,
        False,
        False,
        False,
        False,
        0,
        0,
        False,
        False,
        False,
        False,
        True,
        True,
        True,
        0,
        0,
        False,
    )

    Model.SelectionManager.EnableContourSelection = False
    Model.ClearSelection2(True)
    s = time.perf_counter()
    elapsed = time.perf_counter() - s
    print(f"{__file__} executed in {elapsed:0.2f} seconds until while-loop")
    time.sleep(2)
    while True:
        Model.ViewRotateplusy()
        time.sleep(0.1)

# If running as a script, execute the main function
try:
    if __name__ == "__main__":
        asyncio.run(main())
except KeyboardInterrupt:
    sys.exit()
