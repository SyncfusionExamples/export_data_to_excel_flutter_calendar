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

By using the Workbook and WorkSheet, you can create an excel document based on the appointment data.

```

leading: IconButton(
          icon: const Icon(Icons.import_export),
          onPressed: () async {
            final Workbook workbook = Workbook();
            final Worksheet sheet = workbook.worksheets[0];
            sheet.getRangeByIndex(1, 1).setText('Id');
            sheet.getRangeByIndex(1, 2).setText('Subject');
            sheet.getRangeByIndex(1, 3).setText('Start Time');
            sheet.getRangeByIndex(1, 4).setText('End Time');
            int row = 2;
            for(int i = 0;i<_dataSource.appointments!.length;i++){
              int column = 1;
              final Appointment appointment = _dataSource.appointments![i];
              sheet.getRangeByIndex(row, column).setValue(i+1);
              column++;
              sheet.getRangeByIndex(row, column).setText(appointment.subject);
              column++;
              sheet.getRangeByIndex(row, column).setDateTime(appointment.startTime);
              column++;
              sheet.getRangeByIndex(row, column).setDateTime(appointment.endTime);
              row++;
            }
            final List<int> bytes = workbook.saveAsStream();
            workbook.dispose();

            AnchorElement(
                href:
                "data:application/octet-stream;charset=utf-16le;base64,${base64.encode(bytes)}")
              ..setAttribute("download", "output.xlsx")
              ..click();
          },
        ),

```

## Requirements to run the demo
* [VS Code](https://code.visualstudio.com/download)
* [Flutter SDK v1.22+](https://flutter.dev/docs/development/tools/sdk/overview)
* [For more development tools](https://flutter.dev/docs/development/tools/devtools/overview)

## How to run this application
To run this application, you need to first clone or download the ‘create a flutter maps widget in 10 minutes’ repository and open it in your preferred IDE. Then, build and run your project to view the output.

## Further help
For more help, check the [Syncfusion Flutter documentation](https://help.syncfusion.com/flutter/introduction/overview),
 [Flutter documentation](https://flutter.dev/docs/get-started/install).