import React from 'react';
import { StyleSheet, View, Button } from 'react-native';
// Required to save to cache 
import * as FileSystem from 'expo-file-system';
// ExcelJS
import ExcelJS from 'exceljs';
// Share excel via share dialog
import * as Sharing from 'expo-sharing';
// From @types/node/buffer
import { Buffer as NodeBuffer } from 'buffer';


const styles = StyleSheet.create({
  container: {
    flex: 1,
    backgroundColor: '#fff',
    alignItems: 'center',
    justifyContent: 'center',
  },
});



// This returns a local uri that can be shared
const generateShareableExcel = async (): Promise<string> => {
  const now = new Date();
  const fileName = 'YourFilename.xlsx';
  const fileUri = FileSystem.cacheDirectory + fileName;
  return new Promise<string>((resolve, reject) => {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Me';
    workbook.created = now;
    workbook.modified = now;
    // Add a sheet to work on
    const worksheet = workbook.addWorksheet('My Sheet', {});
    // Just some columns as used on ExcelJS Readme
    worksheet.columns = [
      { header: 'Id', key: 'id', width: 10 },
      { header: 'Name', key: 'name', width: 32 },
      { header: 'D.O.B.', key: 'dob', width: 10, }
    ];
    // Add some test data
    worksheet.addRow({ id: 1, name: 'John Doe', dob: new Date(1970, 1, 1) });
    worksheet.addRow({ id: 2, name: 'Jane Doe', dob: new Date(1969, 2, 3) });

    // Test styling

    // Style first row
    worksheet.getRow(1).font = {
      name: 'Comic Sans MS', family: 4, size: 16, underline: 'double', bold: true
    };
    // Style second column
    worksheet.eachRow((row, rowNumber) => {
      row.getCell(2).font = {
        name: 'Arial Black',
        color: { argb: 'FF00FF00' },
        family: 2,
        size: 14,
        bold: true
      };
    });

    // Write to file
    workbook.xlsx.writeBuffer().then((buffer: ExcelJS.Buffer) => {
      // Do this to use base64 encoding
      const nodeBuffer = NodeBuffer.from(buffer);
      const bufferStr = nodeBuffer.toString('base64');
      FileSystem.writeAsStringAsync(fileUri, bufferStr, {
        encoding: FileSystem.EncodingType.Base64
      }).then(() => {
        resolve(fileUri);
      });
    });
  });
}

const shareExcel = async () => {
  const shareableExcelUri: string = await generateShareableExcel();
  Sharing.shareAsync(shareableExcelUri, {
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // Android
    dialogTitle: 'Your dialog title here', // Android and Web
    UTI: 'com.microsoft.excel.xlsx' // iOS
  }).catch(error => {
    console.error('Error', error);
  }).then(() => {
    console.log('Return from sharing dialog');
  });
}


export default function App() {
  return (
    <View style={styles.container}>
      <Button title='Generate Excel' onPress={shareExcel} />
    </View>
  );
}
