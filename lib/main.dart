import 'dart:convert';
import 'dart:html';

import 'package:flutter/material.dart';
import 'package:flutter/widgets.dart';

import 'package:syncfusion_flutter_calendar/calendar.dart';
import 'package:syncfusion_flutter_xlsio/xlsio.dart';

void main(){
  return runApp(const CalendarToExcel());
}
class CalendarToExcel extends StatelessWidget {
  const CalendarToExcel({Key? key}) : super(key: key);
  @override
  Widget build(BuildContext context) {
    return const MaterialApp(
      home: HomePage(),
    );
  }
}

class HomePage extends StatefulWidget {
  const HomePage({Key? key}) : super(key: key);
  @override
  HomePageState createState() {
    return HomePageState();
  }
}

class HomePageState extends State<HomePage> {
  late AppointmentDataSource _dataSource;
  @override
  void initState() {
    _dataSource = AppointmentDataSource(_getAppointments());
    super.initState();
  }

  List<Appointment> _getAppointments() {
    List<Appointment> appointments = <Appointment>[];
    appointments.add(Appointment(
        startTime: DateTime.now(),
        endTime: DateTime.now().add(const Duration(hours: 2)),
        color: Colors.red,
        subject: 'Meeting'));
    appointments.add(Appointment(
        startTime: DateTime.now().add(const Duration(days: 1)),
        endTime: DateTime.now().add(const Duration(days: 1, hours: 2)),
        color: Colors.red,
        subject: 'Planning'));
    appointments.add(Appointment(
        startTime: DateTime.now().add(const Duration(days: -1)),
        endTime: DateTime.now().add(const Duration(days: -1, hours: 2)),
        color: Colors.red,
        subject: 'Retrospective'));
    appointments.add(Appointment(
        startTime: DateTime.now().add(const Duration(days: 2)),
        endTime: DateTime.now().add(const Duration(hours: 3)),
        color: Colors.red,
        subject: 'Scrum'));
    return appointments;
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('Calendar Demo'),
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


            // final directory = await getExternalStorageDirectory();
            //
            // final path = directory!.path;
            //
            // File file = File('$path/Output.xlsx');
            //
            // await file.writeAsBytes(bytes, flush: true);
            //
            // OpenFile.open('$path/Output.xlsx');
          },
        ),
      ),
      body: SfCalendar(
        dataSource: _dataSource,
      ),
    );
  }
}

class AppointmentDataSource extends CalendarDataSource {
  AppointmentDataSource(List<Appointment> source) {
    appointments = source;
  }
}