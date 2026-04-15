import 'dart:io';
import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:path_provider/path_provider.dart';
import 'package:share_plus/share_plus.dart';
import 'package:archive/archive.dart';

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      debugShowCheckedModeBanner: false,
      home: const HomePage(),
    );
  }
}

class HomePage extends StatefulWidget {
  const HomePage({super.key});

  @override
  State<HomePage> createState() => _HomePageState();
}

class _HomePageState extends State<HomePage> {
  final nameController = TextEditingController();
  final addressController = TextEditingController();
  final cityController = TextEditingController();
  final phoneController = TextEditingController();

  bool isLoading = false;

  Future<void> generateDoc() async {
    if (nameController.text.isEmpty) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('Please enter a name')),
      );
      return;
    }

    setState(() => isLoading = true);

    try {
      // Load template from assets
      final data = await rootBundle.load('assets/template.docx');
      final bytes = data.buffer.asUint8List();

      // Decode the ZIP archive (.docx is a ZIP)
      final archive = ZipDecoder().decodeBytes(bytes);
      final newArchive = Archive();

      for (final file in archive) {
        var content = file.content;

        // Only modify the main document XML
        if (file.name == 'word/document.xml') {
          String xmlString = String.fromCharCodes(content as List<int>);

          // Simple replacement for {{key}}
          // Note: This works best if you type the placeholder in one go in Word.
          final replacements = {
            '{{name}}': nameController.text,
            '{{sender_name}}': addressController.text,
            '{{organization_name}}': cityController.text,
            '{{contact_information}}': phoneController.text,
          };

          replacements.forEach((key, value) {
            xmlString = xmlString.replaceAll(key, value);
          });

          content = xmlString.codeUnits;
        }

        newArchive.addFile(ArchiveFile(file.name, content.length, content));
      }

      final generated = ZipEncoder().encode(newArchive);
      if (generated == null) throw 'Document generation failed';

      // Save to temporary directory
      final dir = await getTemporaryDirectory();
      final timestamp = DateTime.now().millisecondsSinceEpoch;
      final file = File('${dir.path}/generated_$timestamp.docx');

      await file.writeAsBytes(generated);

      // Share file
      await SharePlus.instance.share(
        ShareParams(
          files: [XFile(file.path)],
          subject: 'Generated Document',
        ),
      );

      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          const SnackBar(content: Text('Document generated successfully')),
        );
      }
    } catch (e) {
      if (mounted) {
        debugPrint(e.toString());
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(content: Text('Error: ${e.toString()}')),
        );
      }
    } finally {
      if (mounted) {
        setState(() => isLoading = false);
      }
    }
  }

  Widget buildTextField(String label, TextEditingController controller) {
    return Padding(
      padding: const EdgeInsets.symmetric(vertical: 8),
      child: TextField(
        controller: controller,
        decoration: InputDecoration(
          labelText: label,
          border: OutlineInputBorder(),
        ),
      ),
    );
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: const Text('Generate Word Document')),
      body: Padding(
        padding: const EdgeInsets.all(16),
        child: Column(
          children: [
            buildTextField('Name', nameController),
            buildTextField('Sender', addressController),
            buildTextField('Organization', cityController),
            buildTextField('Contact', phoneController),

            const SizedBox(height: 20),

            isLoading
                ? const CircularProgressIndicator()
                : ElevatedButton(
              onPressed: generateDoc,
              child: const Text('Generate & Download'),
            ),
          ],
        ),
      ),
    );
  }
}
