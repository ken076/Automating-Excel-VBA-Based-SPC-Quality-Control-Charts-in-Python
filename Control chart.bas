X-bar chart
Set rngSample= ActiveSheet.Range("C16:L16")
title = "X-bar Chart"
std=Application.WorksheetFunction.StDev_S(rngSample)
CL=Application.WorksheetFunction.Average(rngSample)
UCL = CL + 1.023 * P5
LCL = CL - 1.023 * P5
rngCL.Value = CL
rngLCL.Value = LCL
rngUCL.Value = UCL

R chart
Set rngSample = ActiveSheet.Range("C17:L17")
title = "R Chart"
std=Application.WorksheetFunction.StDev_S(rngSample)
CL=Application.WorksheetFunction.Average(rngSample)
UCL = 2.574 * P5
LCL = 0 * P5 
rngCL.Value = CL
rngLCL.Value = LCL
rngUCL.Value = UCL
