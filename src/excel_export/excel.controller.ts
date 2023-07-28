/* eslint-disable prettier/prettier */

import { Controller, Get, HttpStatus, Param, Post, Res, UploadedFile, UseInterceptors } from '@nestjs/common';
import { FileInterceptor } from '@nestjs/platform-express';
import { MulterFile } from 'multer';
import * as xlsx from 'xlsx';
import * as path from 'path';
import * as fs from 'fs';


@Controller('excel')
export class ExcelController {
    @Post('upload/:dataFormat/:city')
    @UseInterceptors(FileInterceptor('excel_file'))
    async uploadExcel(@Param('dataFormat') dataFormat: string, @Param('city') city: string, @UploadedFile() file: MulterFile) {
        const fileExtension = path.extname(file?.originalname).toLowerCase();
        if (fileExtension !== '.xls' && fileExtension !== '.xlsx') {
            throw new Error('Invalid file format. Please upload an Excel file.');
        } else {
            const workbook = xlsx.readFile(file.path);
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            let transformedData
            if (dataFormat === 'monthly') {
                transformedData = this.transformData(sheet, city);
            } else {
                transformedData = this.transformDataDaily(sheet, city);
            }

            const newSheet = xlsx.utils.json_to_sheet(transformedData);
            const newWorkbook = xlsx.utils.book_new();
            xlsx.utils.book_append_sheet(newWorkbook, newSheet, 'Sheet1');
            let transformedFileName
            if (city === 'adi') {
                transformedFileName = `uploads/transformed_ahmedabad${file.originalname}`;
                xlsx.writeFile(newWorkbook, transformedFileName);

                return {
                    message: 'File successfully uploaded and modified!',
                    downloadLink: transformedFileName,
                    fileName: `transformed_ahmedabad${file.originalname}`,
                    status: HttpStatus.OK
                };
            } else {
                transformedFileName = `uploads/transformed_junagadh${file.originalname}`;
                xlsx.writeFile(newWorkbook, transformedFileName);

                return {
                    message: 'File successfully uploaded and modified!',
                    downloadLink: transformedFileName,
                    fileName: `transformed_junagadh${file.originalname}`,
                    status: HttpStatus.OK
                };
            }
        }


    }

    private transformData(sheet: xlsx.WorkSheet, city: string) {
        console.log("monthly method called")

        console.log("City: ", city)
        const data = xlsx.utils.sheet_to_json(sheet, { raw: false });
        const transformedData = data.map((row: any) => {
            const dateComponents = row['Date']?.split('/');
            const attendanceDate = new Date(`${dateComponents[2]}/${dateComponents[1]}/${dateComponents[0]}`);
            const inAndOutDateFormat = new Date(`${dateComponents[2]}/${dateComponents[1]}/${dateComponents[0]}`);

            const dateFormatOptions: Intl.DateTimeFormatOptions = { year: 'numeric', month: '2-digit', day: '2-digit' };
            const timeFormatOptions: Intl.DateTimeFormatOptions = { hour: '2-digit', minute: '2-digit', hour12: false };

            const formattedAttendanceDate = new Intl.DateTimeFormat('en-US', dateFormatOptions).format(attendanceDate);

            if (row['First Check In'] && row['Last Check Out']) {
                const inTimeComponents = row['First Check In'].split(':');
                const outTimeComponents = row['Last Check Out'].split(':');

                const inTime = new Date(inAndOutDateFormat);
                inTime.setHours(Number(inTimeComponents[0]), Number(inTimeComponents[1]));

                const outTime = new Date(inAndOutDateFormat);
                outTime.setHours(Number(outTimeComponents[0]), Number(outTimeComponents[1]));

                const lateEntryBoundary = { hour: 10, minute: 30 };
                const earlyExitBoundary = { hour: 18, minute: 30 };

                const isEmptyInTime = inTimeComponents[0] === '' && inTimeComponents[1] === '';
                const isEmptyOutTime = outTimeComponents[0] === '' && outTimeComponents[1] === '';

                const isLateEntry = !isEmptyInTime && (inTime.getHours() > lateEntryBoundary.hour || (inTime.getHours() === lateEntryBoundary.hour && inTime.getMinutes() > lateEntryBoundary.minute));
                const isEarlyExit = !isEmptyOutTime && (outTime.getHours() < earlyExitBoundary.hour || (outTime.getHours() === earlyExitBoundary.hour && outTime.getMinutes() < earlyExitBoundary.minute));

                const formattedDate = new Intl.DateTimeFormat('en-US', dateFormatOptions).format(inAndOutDateFormat);


                const formattedInTime = isEmptyInTime ? '' : new Intl.DateTimeFormat('en-GB', timeFormatOptions).format(inTime);
                const formattedOutTime = isEmptyOutTime ? '' : new Intl.DateTimeFormat('en-GB', timeFormatOptions).format(outTime);
                const employeeNumber = row['Employee ID'].split('/')
                if (city === 'adi') {
                    return {
                        ID: '',
                        Series: 'HR-ATT-.YYYY.-',
                        Employee: `TRS/ADI/${employeeNumber[2]}`,
                        Status: 'Present',
                        'Attendance Date': formattedAttendanceDate,
                        Company: 'Techrover Solutions Pvt Ltd',
                        'Employee Name': row['First Name'],
                        Shift: 'All_Day',
                        'In Time': `${formattedDate} ${formattedInTime}`,
                        'Out Time': `${formattedDate} ${formattedOutTime}`,
                        'Late Entry': isLateEntry ? 'YES' : 'NO',
                        'Early Exit': isEarlyExit ? 'YES' : 'NO',
                    };
                } else {
                    return {
                        ID: '',
                        Series: 'HR-ATT-.YYYY.-',
                        Employee: `TRS/JND/${employeeNumber[2]}`,
                        Status: 'Present',
                        'Attendance Date': formattedAttendanceDate,
                        Company: 'Techrover Solutions Pvt Ltd',
                        'Employee Name': row['First Name'],
                        Shift: 'All_Day',
                        'In Time': `${formattedDate} ${formattedInTime}`,
                        'Out Time': `${formattedDate} ${formattedOutTime}`,
                        'Late Entry': isLateEntry ? 'YES' : 'NO',
                        'Early Exit': isEarlyExit ? 'YES' : 'NO',
                    };
                }

            }
            else {
                let inTimeComponents;
                let outTimeComponents;

                if (!row['First Check In']) {
                    inTimeComponents = ['', ''];
                } else {
                    inTimeComponents = row['First Check In'].split(':');
                }

                if (!row['Last Check Out']) {
                    outTimeComponents = ['', ''];
                } else {
                    outTimeComponents = row['Last Check Out'].split(':');
                }

                const lateEntryBoundary = { hour: 10, minute: 30 };
                const earlyExitBoundary = { hour: 18, minute: 30 };

                const isEmptyInTime = inTimeComponents[0] === '' && inTimeComponents[1] === '';
                const isEmptyOutTime = outTimeComponents[0] === '' && outTimeComponents[1] === '';

                const inTime = new Date(inAndOutDateFormat);
                inTime.setHours(isEmptyInTime ? 0 : Number(inTimeComponents[0]), isEmptyInTime ? 0 : Number(inTimeComponents[1]));

                const outTime = new Date(inAndOutDateFormat);
                outTime.setHours(isEmptyOutTime ? 0 : Number(outTimeComponents[0]), isEmptyOutTime ? 0 : Number(outTimeComponents[1]));

                const isLateEntry = !isEmptyInTime && (inTime.getHours() > lateEntryBoundary.hour || (inTime.getHours() === lateEntryBoundary.hour && inTime.getMinutes() > lateEntryBoundary.minute));
                const isEarlyExit = !isEmptyOutTime && (outTime.getHours() < earlyExitBoundary.hour || (outTime.getHours() === earlyExitBoundary.hour && outTime.getMinutes() < earlyExitBoundary.minute));

                const formattedDate = new Intl.DateTimeFormat('en-US', dateFormatOptions).format(inAndOutDateFormat);


                const formattedInTime = isEmptyInTime ? '' : new Intl.DateTimeFormat('en-GB', timeFormatOptions).format(inTime);
                const formattedOutTime = isEmptyOutTime ? '' : new Intl.DateTimeFormat('en-GB', timeFormatOptions).format(outTime);

                const employeeNumber = row['Employee ID'].split('/')
                if (city === 'adi') {
                    return {
                        ID: '',
                        Series: 'HR-ATT-.YYYY.-',
                        Employee: `TRS/ADI/${employeeNumber[2]}`,
                        Status: 'Present',
                        'Attendance Date': formattedAttendanceDate,
                        Company: 'Techrover Solutions Pvt Ltd',
                        'Employee Name': row['First Name'],
                        Shift: 'All_Day',
                        'In Time': `${formattedDate} ${formattedInTime}`,
                        'Out Time': `${formattedDate} ${formattedOutTime}`,
                        'Late Entry': isLateEntry ? 'YES' : 'NO',
                        'Early Exit': isEarlyExit ? 'YES' : 'NO',
                    };
                } else {
                    return {
                        ID: '',
                        Series: 'HR-ATT-.YYYY.-',
                        Employee: `TRS/JND/${employeeNumber[2]}`,
                        Status: 'Present',
                        'Attendance Date': formattedAttendanceDate,
                        Company: 'Techrover Solutions Pvt Ltd',
                        'Employee Name': row['First Name'],
                        Shift: 'All_Day',
                        'In Time': `${formattedDate} ${formattedInTime}`,
                        'Out Time': `${formattedDate} ${formattedOutTime}`,
                        'Late Entry': isLateEntry ? 'YES' : 'NO',
                        'Early Exit': isEarlyExit ? 'YES' : 'NO',
                    };
                }
            }

        });

        return transformedData;
    }

    private transformDataDaily(sheet: xlsx.WorkSheet, city: string) {
        console.log("Daily method called")
        console.log("City: ", city)
        let data
        if (city === 'adi') {
            delete sheet['A1'];
            delete sheet['A2'];

            const range = xlsx.utils.decode_range(sheet['!ref']);
            for (let row = range.s.r + 2; row <= range.e.r; row++) {
                for (let col = range.s.c; col <= range.e.c; col++) {
                    const fromCell = xlsx.utils.encode_cell({ r: row, c: col });
                    const toCell = xlsx.utils.encode_cell({ r: row - 2, c: col });
                    sheet[toCell] = sheet[fromCell];
                    delete sheet[fromCell];
                }
            }

            sheet['!ref'] = xlsx.utils.encode_range({ s: { r: range.s.r, c: range.s.c }, e: { r: range.e.r - 2, c: range.e.c } });
            data = xlsx.utils.sheet_to_json(sheet, { raw: false });
            console.log("Data: ", data);
        } else {
            delete sheet['A1'];
            delete sheet['A2'];
            delete sheet['A3'];
            delete sheet['A4'];
            delete sheet['G3'];

            const range = xlsx.utils.decode_range(sheet['!ref']);
            for (let row = range.s.r + 2; row <= range.e.r; row++) {
                for (let col = range.s.c; col <= range.e.c; col++) {
                    const fromCell = xlsx.utils.encode_cell({ r: row, c: col });
                    const toCell = xlsx.utils.encode_cell({ r: row - 4, c: col });
                    sheet[toCell] = sheet[fromCell];
                    delete sheet[fromCell];
                }
            }

            sheet['!ref'] = xlsx.utils.encode_range({ s: { r: range.s.r, c: range.s.c }, e: { r: range.e.r - 2, c: range.e.c } });
            data = xlsx.utils.sheet_to_json(sheet, { raw: false });
            console.log("Data: ", data);
        }

        const transformedData = data.map((row: any) => {
            // console.log([row['Employee ID']])
            const dateComponents = row['Date']?.split('-');
            const attendanceDate = new Date(`${dateComponents[2]}/${dateComponents[1]}/${dateComponents[0]}`);
            const inAndOutDateFormat = new Date(`${dateComponents[2]}/${dateComponents[1]}/${dateComponents[0]}`);

            const dateFormatOptions: Intl.DateTimeFormatOptions = { year: 'numeric', month: '2-digit', day: '2-digit' };
            const timeFormatOptions: Intl.DateTimeFormatOptions = { hour: '2-digit', minute: '2-digit', hour12: false };

            const formattedAttendanceDate = new Intl.DateTimeFormat('en-US', dateFormatOptions).format(attendanceDate);

            if (row['First Punch'] && row['Last Punch']) {
                const inTimeComponents = row['First Punch'].split(':');
                const outTimeComponents = row['Last Punch'].split(':');

                const inTime = new Date(inAndOutDateFormat);
                inTime.setHours(Number(inTimeComponents[0]), Number(inTimeComponents[1]));

                const outTime = new Date(inAndOutDateFormat);
                outTime.setHours(Number(outTimeComponents[0]), Number(outTimeComponents[1]));

                const lateEntryBoundary = { hour: 10, minute: 30 };
                const earlyExitBoundary = { hour: 18, minute: 30 };

                const isEmptyInTime = inTimeComponents[0] === '' && inTimeComponents[1] === '';
                const isEmptyOutTime = outTimeComponents[0] === '' && outTimeComponents[1] === '';

                const isLateEntry = !isEmptyInTime && (inTime.getHours() > lateEntryBoundary.hour || (inTime.getHours() === lateEntryBoundary.hour && inTime.getMinutes() > lateEntryBoundary.minute));
                const isEarlyExit = !isEmptyOutTime && (outTime.getHours() < earlyExitBoundary.hour || (outTime.getHours() === earlyExitBoundary.hour && outTime.getMinutes() < earlyExitBoundary.minute));

                const formattedDate = new Intl.DateTimeFormat('en-US', dateFormatOptions).format(inAndOutDateFormat);

                const formattedInTime = isEmptyInTime ? '' : new Intl.DateTimeFormat('en-GB', timeFormatOptions).format(inTime);
                const formattedOutTime = isEmptyOutTime ? '' : new Intl.DateTimeFormat('en-GB', timeFormatOptions).format(outTime);
                let employeeNumber: any = []
                if (row['Employee ID'].length > 1) {
                    employeeNumber = row['Employee ID']
                } else {
                    employeeNumber = '0' + row['Employee ID'][0];
                }
                if (city === 'adi') {
                    return {
                        ID: '',
                        Series: 'HR-ATT-.YYYY.-',
                        Employee: `TRS/ADI/${employeeNumber}`,
                        Status: 'Present',
                        'Attendance Date': formattedAttendanceDate,
                        Company: 'Techrover Solutions Pvt Ltd',
                        'Employee Name': row['First Name'],
                        Shift: 'All_Day',
                        'In Time': `${formattedDate} ${formattedInTime}`,
                        'Out Time': `${formattedDate} ${formattedOutTime}`,
                        'Late Entry': isLateEntry ? 'YES' : 'NO',
                        'Early Exit': isEarlyExit ? 'YES' : 'NO',
                    };
                } else {
                    return {
                        ID: '',
                        Series: 'HR-ATT-.YYYY.-',
                        Employee: `TRS/JND/${employeeNumber}`,
                        Status: 'Present',
                        'Attendance Date': formattedAttendanceDate,
                        Company: 'Techrover Solutions Pvt Ltd',
                        'Employee Name': row['First Name'],
                        Shift: 'All_Day',
                        'In Time': `${formattedDate} ${formattedInTime}`,
                        'Out Time': `${formattedDate} ${formattedOutTime}`,
                        'Late Entry': isLateEntry ? 'YES' : 'NO',
                        'Early Exit': isEarlyExit ? 'YES' : 'NO',
                    };
                }

            }
            else {
                let inTimeComponents;
                let outTimeComponents;

                if (!row['First Punch']) {
                    inTimeComponents = ['', ''];
                } else {
                    inTimeComponents = row['First Punch'].split(':');
                }

                if (!row['Last Punch']) {
                    outTimeComponents = ['', ''];
                } else {
                    outTimeComponents = row['Last Punch'].split(':');
                }

                const lateEntryBoundary = { hour: 10, minute: 30 };
                const earlyExitBoundary = { hour: 18, minute: 30 };

                const isEmptyInTime = inTimeComponents[0] === '' && inTimeComponents[1] === '';
                const isEmptyOutTime = outTimeComponents[0] === '' && outTimeComponents[1] === '';

                const inTime = new Date(inAndOutDateFormat);
                inTime.setHours(isEmptyInTime ? 0 : Number(inTimeComponents[0]), isEmptyInTime ? 0 : Number(inTimeComponents[1]));

                const outTime = new Date(inAndOutDateFormat);
                outTime.setHours(isEmptyOutTime ? 0 : Number(outTimeComponents[0]), isEmptyOutTime ? 0 : Number(outTimeComponents[1]));

                const isLateEntry = !isEmptyInTime && (inTime.getHours() > lateEntryBoundary.hour || (inTime.getHours() === lateEntryBoundary.hour && inTime.getMinutes() > lateEntryBoundary.minute));
                const isEarlyExit = !isEmptyOutTime && (outTime.getHours() < earlyExitBoundary.hour || (outTime.getHours() === earlyExitBoundary.hour && outTime.getMinutes() < earlyExitBoundary.minute));

                const formattedDate = new Intl.DateTimeFormat('en-US', dateFormatOptions).format(inAndOutDateFormat);

                const formattedInTime = isEmptyInTime ? '' : new Intl.DateTimeFormat('en-GB', timeFormatOptions).format(inTime);
                const formattedOutTime = isEmptyOutTime ? '' : new Intl.DateTimeFormat('en-GB', timeFormatOptions).format(outTime);
                let employeeNumber: any = []
                if (row['Employee ID'].length > 1) {
                    employeeNumber = row['Employee ID']
                } else {
                    employeeNumber = '0' + row['Employee ID'][0];
                }
                if (city === 'adi') {
                    return {
                        ID: '',
                        Series: 'HR-ATT-.YYYY.-',
                        Employee: `TRS/ADI/${employeeNumber}`,
                        Status: 'Present',
                        'Attendance Date': formattedAttendanceDate,
                        Company: 'Techrover Solutions Pvt Ltd',
                        'Employee Name': row['First Name'],
                        Shift: 'All_Day',
                        'In Time': `${formattedDate} ${formattedInTime}`,
                        'Out Time': `${formattedDate} ${formattedOutTime}`,
                        'Late Entry': isLateEntry ? 'YES' : 'NO',
                        'Early Exit': isEarlyExit ? 'YES' : 'NO',
                    };
                } else {
                    return {
                        ID: '',
                        Series: 'HR-ATT-.YYYY.-',
                        Employee: `TRS/JND/${employeeNumber}`,
                        Status: 'Present',
                        'Attendance Date': formattedAttendanceDate,
                        Company: 'Techrover Solutions Pvt Ltd',
                        'Employee Name': row['First Name'],
                        Shift: 'All_Day',
                        'In Time': `${formattedDate} ${formattedInTime}`,
                        'Out Time': `${formattedDate} ${formattedOutTime}`,
                        'Late Entry': isLateEntry ? 'YES' : 'NO',
                        'Early Exit': isEarlyExit ? 'YES' : 'NO',
                    };
                }
            }
        });
        return transformedData;
    }

    @Get('download/:fileName')
    async downloadTransformedExcel(@Param('fileName') fileName: string, @Res() res) {
        const file = `uploads/${fileName}`;
        const stat = fs.statSync(file);

        res.set({
            'Content-Type': 'application/octet-stream',
            'Content-Length': stat.size,
        });
        fs.createReadStream(file).pipe(res);
    }
}


