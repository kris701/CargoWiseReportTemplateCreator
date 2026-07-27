<p align="center">
    <img src="https://github.com/user-attachments/assets/7f3bd98c-97b5-471e-a3bb-75dba11cb3b5" width="200" height="200" />
</p>

[![Build and Publish](https://github.com/kris701/CargoWiseReportTemplateCreator/actions/workflows/dotnet.yml/badge.svg)](https://github.com/kris701/CargoWiseReportTemplateCreator/actions/workflows/dotnet.yml)
![Nuget](https://img.shields.io/nuget/v/CargoWiseReportTemplateCreator)
![Nuget](https://img.shields.io/nuget/dt/CargoWiseReportTemplateCreator)
![GitHub last commit (branch)](https://img.shields.io/github/last-commit/kris701/CargoWiseReportTemplateCreator/main)
![GitHub commit activity (branch)](https://img.shields.io/github/commit-activity/m/kris701/CargoWiseReportTemplateCreator)
![Static Badge](https://img.shields.io/badge/Platform-Windows-blue)
![Static Badge](https://img.shields.io/badge/Platform-Linux-blue)
![Static Badge](https://img.shields.io/badge/Framework-dotnet--10.0-green)

# CargoWise Report Template Creator

This is a simple little tool to create report templates for CargoWise.
It is packaged on the [NuGet Package Manager](https://www.nuget.org/packages/CargoWiseReportTemplateCreator/) as a dotnet tool, so you can install it by writing `dotnet tool install CargoWiseReportTemplateCreator` into a terminal.
You can then use the tool by writing `cwreporttemplatecreator` in a terminal.

You can use the tool multiple times on the same report file.
The tool will simply create a new sheet for whatever table you
are adding to the report. So as an example, you can do the following:

```powershell
cwreporttemplatecreator -t JobCharge -c JR_PK JR_JH JR_GE -o "Replication Data Template.xls"
cwreporttemplatecreator -t AccChargeCode -c AC_PK AC_Code AC_Desc -o "Replication Data Template.xls"
```

This will create a single file `Replication Data Template.xls` that 
has two sheets, one for JobCharge and one for AccChargeCode.
You can do this with as many tables and columns you want.

Afterwards, all you have to do in CargoWise is go to Reports and make 
a new report with this template.
