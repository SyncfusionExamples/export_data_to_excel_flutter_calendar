# How to export the Flutter calendar appointments into Excel?

This example demonstrates how to export the Flutter calendar appointments into Excel.

## Install the required packages.

Using the [Syncfusion xlsio](https://www.syncfusion.com/flutter-widgets/excel-library) package, the Flutter Calendar's appointments can be exported to an Excel sheet.

```
  syncfusion_flutter_calendar: ^20.3.56
  intl: ^0.17.0
  syncfusion_flutter_xlsio: ^20.3.56-beta
  path_provider: ^2.0.11
  open_file: ^3.2.1

```

## Creating an excel document based on the appointments data

By using the Workbook and WorkSheet, you can create an excel document based on the appointment data. Here by using the getRangeByIndex from the WorkSheet the calendar appointment data have been stored in an excel document.

## Requirements to run the demo
* [VS Code](https://code.visualstudio.com/download)
* [Flutter SDK v1.22+](https://flutter.dev/docs/development/tools/sdk/overview)
* [For more development tools](https://flutter.dev/docs/development/tools/devtools/overview)

## How to run this application
To run this application, you need to first clone or download the ‘create a flutter maps widget in 10 minutes’ repository and open it in your preferred IDE. Then, build and run your project to view the output.

## Further help
For more help, check the [Syncfusion Flutter documentation](https://help.syncfusion.com/flutter/introduction/overview),
 [Flutter documentation](https://flutter.dev/docs/get-started/install).