import 'dart:io';

import 'package:clipboard/clipboard.dart';
import 'package:excel/excel.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:flutter_convert_xlsx/progress_dialog.dart';
import 'package:html_unescape/html_unescape.dart';
import 'package:xml/xml.dart';
import 'package:oktoast/oktoast.dart';

void main() async {
  WidgetsFlutterBinding.ensureInitialized();
  runApp(MyApp());
  // runApp(FilePickerDemo());
}

class MyApp extends StatelessWidget {
  // This widget is the root of your application.
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Flutter converter',
      debugShowCheckedModeBanner: false,
      theme: ThemeData(
        // This is the theme of your application.
        //
        // Try running your application with "flutter run". You'll see the
        // application has a blue toolbar. Then, without quitting the app, try
        // changing the primarySwatch below to Colors.green and then invoke
        // "hot reload" (press "r" in the console where you ran "flutter run",
        // or simply save your changes to "hot reload" in a Flutter IDE).
        // Notice that the counter didn't reset back to zero; the application
        // is not restarted.
        primarySwatch: Colors.blue,
      ),
      home: MyHomePage(title: 'Flutter convert .xlsx to .xml'),
      builder: (_, Widget? child) => OKToast(child: child!),
    );
  }
}

class MyHomePage extends StatefulWidget {
  MyHomePage({Key? key, required this.title}) : super(key: key);

  // This widget is the home page of your application. It is stateful, meaning
  // that it has a State object (defined below) that contains fields that affect
  // how it looks.

  // This class is the configuration for the state. It holds the values (in this
  // case the title) provided by the parent (in this case the App widget) and
  // used by the build method of the State. Fields in a Widget subclass are
  // always marked "final".

  final String title;

  @override
  _MyHomePageState createState() => _MyHomePageState();
}

class _MyHomePageState extends State<MyHomePage> {
  ProgressDialog? pr;

  final builderSlot1 = XmlBuilder();
  final builderSlot2 = XmlBuilder();

  final textEditController1 = TextEditingController();
  final textEditController2 = TextEditingController();

  bool success = false;
  bool isLog = true;

  @override
  void initState() {
    super.initState();
    pr = ProgressDialog(context,
        type: ProgressDialogType.Normal, isDismissible: false, showLogs: true);
    pr?.style(
        message: 'Downloading file...',
        borderRadius: 10.0,
        backgroundColor: Colors.white,
        progressWidget: CircularProgressIndicator(),
        elevation: 10.0,
        insetAnimCurve: Curves.easeInOut,
        progress: 0.0,
        maxProgress: 100.0,
        progressTextStyle:
            TextStyle(color: Colors.black, fontSize: 13.0, fontWeight: FontWeight.w400),
        messageTextStyle:
            TextStyle(color: Colors.black, fontSize: 19.0, fontWeight: FontWeight.w600));
  }

  void _filePicker() async {
    setState(() {
      success = false;
    });
    final result = await FilePicker.platform.pickFiles(type: FileType.any, allowMultiple: false);
    await pr?.show();

    if (result?.files.first != null) {
      var fileBytes = result?.files.first.bytes;

      var excel = Excel.decodeBytes(fileBytes!);

      _convertToXml(builderSlot1, excel);
      _convertToXml(builderSlot2, excel, slot1: false);

      setState(() {
        success = true;
      });
    }
    pr?.hide();

    textEditController1.text = builderSlot1.buildDocument().toXmlString(pretty: true);
    textEditController2.text = builderSlot2.buildDocument().toXmlString(pretty: true);
  }

  void _convertToXml(XmlBuilder builder, Excel excel, {bool slot1 = true}) {
    builder.processing('xml', 'version="1.0"');

    int column = slot1 ? 2 : 3;
    builder.element('resources', nest: () {
      for (var table in excel.tables.keys) {
        var index = 0;
        try {
          for (var row in excel.tables[table]!.rows) {
            if (index > 0 && row[0]?.value != null && row[column]?.value != null) {
              builder.element('string', nest: () {
                builder.attribute('name',
                    '${table.toLowerCase().replaceAll(" ", "_")}_${row[0]?.value.toLowerCase().replaceAll(" ", "_")}');
                String? value = row[column]?.value?.toString();
                builder.text(value!.replaceAll('-', '&#8211;').replaceAll('\'', '\\'''));
              });
            }
            index += 1;
          }
        } catch (e) {
          e.toString();
        }
      }
    });
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: Text(widget.title),
        actions: [
          Center(child: Text("Log")),
          Switch(
            value: isLog,
            onChanged: (value) {
              setState(() {
                isLog = value;
              });
            },
            activeTrackColor: Colors.lightBlueAccent,
            activeColor: Colors.blue,
          ),
        ],
      ),
      body: Center(
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.stretch,
          children: <Widget>[
            RaisedButton(
              onPressed: _filePicker,
              child: Container(
                padding: EdgeInsets.all(20),
                child: Text(
                  'Upload file to convert...',
                  style: Theme.of(context).textTheme.headline4!.copyWith(color: Colors.black),
                ),
              ),
            ),
            SizedBox(
              height: 20,
            ),
            Container(
              height: 200,
              child: Row(
                crossAxisAlignment: CrossAxisAlignment.stretch,
                children: [
                  Spacer(),
                  Visibility(
                    maintainSize: true,
                    maintainAnimation: true,
                    maintainState: true,
                    visible: success,
                    child: RaisedButton(
                      color: Colors.blue,
                      onPressed: () {
                        _clipboardCopy(textEditController1.text);
                      },
                      child: Container(
                        padding: EdgeInsets.all(20),
                        child: Text(
                          'Copy Xml EN',
                          style: TextStyle(color: Colors.white, fontSize: 20),
                        ),
                      ),
                    ),
                  ),
                  SizedBox(
                    width: 50,
                  ),
                  Visibility(
                    maintainSize: true,
                    maintainAnimation: true,
                    maintainState: true,
                    visible: success,
                    child: RaisedButton(
                      color: Colors.blue,
                      onPressed: () {
                        _clipboardCopy(textEditController2.text);
                      },
                      child: Container(
                        padding: EdgeInsets.all(20),
                        child: Text(
                          'Copy Xml TH',
                          style: TextStyle(color: Colors.white, fontSize: 20),
                        ),
                      ),
                    ),
                  ),
                  Spacer(),
                ],
              ),
            ),
            if (isLog)
              Expanded(
                child: Visibility(
                  maintainSize: true,
                  maintainAnimation: true,
                  maintainState: true,
                  visible: success,
                  child: SingleChildScrollView(
                    child: Container(
                      padding: EdgeInsets.all(20),
                      child: Row(
                        crossAxisAlignment: CrossAxisAlignment.start,
                        children: [
                          Expanded(child: Text(textEditController1.text)),
                          SizedBox(width: 20),
                          Expanded(child: Text(textEditController2.text)),
                        ],
                      ),
                    ),
                  ),
                ),
              ),
            if (!isLog) Spacer()
          ],
        ),
      ),
    );
  }

  void _clipboardCopy(String value) {
    FlutterClipboard.copy(value).then((value) {
      showToast(
        'copied...',
        position: ToastPosition.bottom,
        backgroundColor: Colors.black.withOpacity(0.8),
        radius: 12.0,
        textStyle: Theme.of(context).textTheme.headline3!.merge(TextStyle(color: Colors.white)),
        animationBuilder: const Miui10AnimBuilder(),
      );
    });
  }
}
