import express from "express";
import cors from "cors";
import mongoose from "mongoose";
import "dotenv/config";
import ExcelJS from "exceljs";
import { Parser } from 'json2csv';
import PDFDocument from "pdfkit"; 
import fs from "fs";
import path from 'path';
import pdf from "pdfkit"
import os from 'os';
import { format, eachDayOfInterval } from 'date-fns';
// import { format } from 'date-fns';
// import { toZonedTime } from 'date-fns-tz'

const port = 3000;
const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(cors({ origin: "*" }));
const username = process.env.MONGO_USERNAME;
const password = encodeURIComponent(process.env.MONGO_PASSWORD);

// const Musername = process.env.MONGO_USERNAME;
// const Mpassword = process.env.MONGO_PASSWORD;

// const dbConnection = await mongoose.connect(

// //   "mongodb+srv://" +
// //     Musername +
// //     ":" +
// //     Mpassword +
// //     "@cluster0.4ont6qs.mongodb.net/attendance?retryWrites=true&w=majority&appName=Cluster0"
// );
// if (dbConnection)
//   app.listen(port, () => console.log("Server Started on port " + port));
//..............................................................................

mongoose.connect(
    "mongodb+srv://" +
    username +
    ":" +
    password +
    "@cluster0.3j0ywmp.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0/attendance"

)
    .then(() => {
        console.log("MongoDB connected");
        app.listen(port, () => {
            console.log(`Server running on port ${port}`);
        });
    })
    .catch((error) => console.error("MongoDB connection error:", error));


// mongoose
//   .connect("mongodb://127.0.0.1:27017/attendance")
//   .then(() =>
//     app.listen(port, () => console.log("server started at port " + port))
//   )
//   .catch((err) => console.log("Failed to connect to MongoDB", err));


const timeZone = "Asia/Kolkata";
const date = new Date();
// const zonedDated = toZonedTime(date, timeZone)
const formattedDate = format(date, "dd-MM-yyyy HH:mm:ss", { timeZone });
let d = new Date().toLocaleString("en-US", { timeZone: "Asia/Kolkata" });

const attendanceSchema = new mongoose.Schema({
    date: {
        type: Date,
        default: d,
    },
    faculty: {
        type: String,
        required: true,
    },
    attendance: {
        type: Object,
        required: true,
    },
});
const facultySchema = new mongoose.Schema({

    facultyList: [{
        id: {
            type: String,
            required: true,
            // default: Date.now(),
            default: () => new mongoose.Types.ObjectId().toString(),
        },
        name:
        {
            type: String,
            required: true,
        }

    }]
});

// aadharCard: {
//     type: String,
//     required: true,
// },
// name: {
//     type: String,  
//     required: true,
// },
const studentSchema = new mongoose.Schema({
    id: {
        type: Number,
        default: Date.now(),
    },

    faculty: {
        type: String,
        required: true,
    },

    students: {
        type: Array,
        required: true,
    }

});

const attendanceModel = mongoose.model("record", attendanceSchema, "record");
const facultyModel = mongoose.model("faculty", facultySchema, "faculty");
const studentModel = mongoose.model("student", studentSchema, "student");


app.post("/saveAttendance", (req, res) => {
    const dataToSave = new attendanceModel(req.body);
    dataToSave
        .save()
        .then((response) => res.send(response))
        .catch((error) => {
            console.log(error);
            res.send(false);
        });
});

app.get('/getAttendanceFaculties', async (req, res) => {
    try {
        const faculties = await attendanceModel.distinct('faculty');
        res.json(faculties);
    } catch (error) {
        console.error('Error fetching faculties:', error);
        res.status(500).send('Error fetching faculties');
    }
});

app.get("/getFaculty", async (req, res) => {
    const facultyListt = await facultyModel.find({}).select("-__v");
    if (facultyListt) res.json(facultyListt);
    else res.json(false);
});


app.post("/saveFaculty", async (req, res) => {
    const dataToSave = new facultyModel(req.body);
    console.log(dataToSave, "dataToSave")
    dataToSave
        .save()
        .then((response) => {
            console.log("Save response:", response);
            res.json(response)
        })
        .catch((error) => {
            console.log("error", error);
            res.send(false);
        });
});

app.get("/getStudent", async (req, res) => {
    const { faculty } = req.query;
    try {
        let studentList;
        if (faculty) {
            studentList = await studentModel.find({ faculty }).select("-__v");
        } else {
            studentList = await studentModel.find({}).select("-__v");
        }
        res.json(studentList);
    } catch (error) {
        console.error(error);
        res.status(500).json({ error: "error" });
    }
});

app.post("/saveStudent", async (req, res) => {
    // const { name, faculty, aadhaarCard } = req.body
    const dataToSave = new studentModel(req.body);
    console.log("dataToSave", dataToSave);

    dataToSave
        .save()
        .then((response) => res.json(response))
        .catch((error) => {
            console.log(error);
            res.send(false);

        });
});

// const idToDelete = req.params.id;
// const deletedStudent = await studentModel.findOneAndDelete({
//     id: idToDelete
// });
// if (deletedStudent) res.json("Student Deleted");
// else res.json(false);

app.delete("/deleteStudent/:id", async (req, res) => {
    const idDelete = req.params.id;
    console.log(idDelete, "idDelete")
    try {
        const fResult = await studentModel.deleteOne({
            "students.id": idDelete

        });
        console.log(fResult, "fResult")
        if (fResult) {
            console.log("student delete", fResult);
            res.json("student deleted")
        }
        else {
            res.status(400).json("error")
        }
    }
    catch (error) {
        console.log(error)
        res.status(500).json({ error: "error" })
    };
});


app.delete("/deleteFaculty/:id", async (req, res) => {
    const facultyIdToDelete = req.params.id;

    try {
        const documents = await facultyModel.find({ "facultyList.id": facultyIdToDelete });

        if (!documents.length) {
            return res.json(false);
        }

        let modifiedCount = 0;

        for (const doc of documents) {
            const updatedFacultyList = doc.facultyList.filter(faculty => faculty.id !== facultyIdToDelete);
            const result = await facultyModel.updateOne(
                { _id: doc._id },
                { $set: { facultyList: updatedFacultyList } }
            );
            if (result.modifiedCount > 0) {
                modifiedCount++;
            }
        }

        if (modifiedCount > 0) {
            res.json("Faculty Deleted");
        } else {
            res.json(false);
        }
    } catch (error) {
        console.log("Error deleting faculty:", error);
        res.json(false);
    }
});


app.delete("/deleteFacultyMany", async (req, res) => {
    const facultyIdsToDelete = req.body.ids;
    try {
        const documents = await facultyModel.find({ "facultyList.id": { $in: facultyIdsToDelete } });

        if (!documents.length) {
            return res.json(false);
        }
        let modifiedCount = 0;
        for (const doc of documents) {
            const updatedFacultyList = doc.facultyList.filter(faculty => !facultyIdsToDelete.includes(faculty.id));
            const result = await facultyModel.updateOne(
                { _id: doc._id },
                { $set: { facultyList: updatedFacultyList } }
            );
            if (result.modifiedCount > 0) {
                modifiedCount++;
            }
        }
        if (modifiedCount > 0) {
            res.json("Faculty Deleted");
        } else {
            res.json(false);
        }
    } catch (error) {
        console.log("Error deleting faculty:", error);
        res.json(false);
    }
});

// exportAttendance
app.get("/exportAttendance/:startDate/:endDate", async (req, res) => {
    const { startDate, endDate } = req.params;
    const { faculties } = req.query; 

    try {
        const endOfDay = new Date(endDate);
        endOfDay.setHours(23, 59, 59, 999);

        let filter = {
            date: {
                $gte: new Date(startDate),
                $lte: endOfDay
            }
        };

        if (faculties) {
            const facultyList = faculties.split(",");
            filter.faculty = { $in: facultyList };
        }

        const attendanceData = await attendanceModel.find(filter);

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Attendance');

        const formattedStartDate = format(new Date(startDate), 'dd-MM-yyyy', { timeZone: 'Asia/Kolkata' });
        const formattedEndDate = format(new Date(endDate), 'dd-MM-yyyy', { timeZone: 'Asia/Kolkata' });

        const dateRange = eachDayOfInterval({ start: new Date(startDate), end: new Date(endOfDay) }).map(date => format(date, 'dd-MM-yyyy'));
        const headerRow = ['Faculty', 'Student Name', 'Date', ...dateRange];
        const headerRowCell = worksheet.addRow(headerRow);
        headerRowCell.height = 28;
        headerRowCell.eachCell((cell) => {
            cell.font = { bold: true };
            cell.alignment = { vertical: 'center', horizontal: 'center' };
        });

        worksheet.getColumn(1).width = 16;
        worksheet.getColumn(2).width = 16;
        worksheet.getColumn(3).width = 12;

        worksheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
            row.eachCell({ includeEmpty: true }, function (cell, colNumber) {
                if (colNumber === 1 || colNumber === 2 || colNumber === 3) {
                    cell.alignment = { horizontal: 'center' };
                }
            });
        });

        for (let i = 4; i < dateRange.length + 4; i++) {
            const column = worksheet.getColumn(i);
            column.width = 12;
            column.alignment = { horizontal: 'center' };
        }

        const attendanceMap = {};
        attendanceData.forEach(record => {
            const { date, faculty, attendance } = record;
            const formattedDate = format(date, 'dd-MM-yyyy');

            if (!attendanceMap[faculty]) {
                attendanceMap[faculty] = {};
            }

            for (const studentId in attendance) {
                const studentData = attendance[studentId];
                if (!attendanceMap[faculty][studentData.name]) {
                    attendanceMap[faculty][studentData.name] = {};
                }
                attendanceMap[faculty][studentData.name][formattedDate] = studentData.status;
            }
        });

        Object.keys(attendanceMap).forEach(faculty => {
            const students = attendanceMap[faculty];
            let isFirstStudent = true;

            Object.keys(students).forEach(studentName => {
                const attendanceByDate = students[studentName];
                const row = [isFirstStudent ? faculty : '', studentName, ''];

                dateRange.forEach(date => {
                    row.push(attendanceByDate[date] || 'A');
                });

                const newRow = worksheet.addRow(row);
                newRow.eachCell((cell) => {
                    cell.alignment = { horizontal: 'center' };
                });

                dateRange.forEach((date, index) => {
                    const cell = newRow.getCell(index + 4);
                    if (cell.value === 'P') {
                        cell.font = {
                            color: { argb: 'FF005500' },
                            bold: true
                        };
                    } else if (cell.value === 'A') {
                        cell.font = { color: { argb: 'FFFF0000' } };
                    }
                });

                isFirstStudent = false;
            });
        });

        const filename = `attendance_${startDate}_${endDate}.xlsx`;
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        await workbook.xlsx.write(res);

        res.end();
    } catch (error) {
        console.error('Error exporting attendance:', error);
        res.status(500).send('Internal Server Error');
    }
});

// exportAttendancePDF
app.get('/exportAttendancePDF/:startDate/:endDate', async (req, res) => {
    const { startDate, endDate } = req.params;
    const { faculties } = req.query;

    try {
        const endOfDay = new Date(endDate);
        endOfDay.setHours(23, 59, 59, 999);

        const attendanceData = await attendanceModel.find({
            date: {
                $gte: new Date(startDate),
                $lte: endOfDay
            },
            faculty: { $in: faculties.split(',') } 
        });

        const doc = new pdf();
        const filename = `attendance_${startDate}_${endDate}.pdf`;
        res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
        res.setHeader("Content-Type", "application/pdf");
        doc.pipe(res);

        const startDateFormatted = formatDate(new Date(startDate));
        const endDateFormatted = formatDate(new Date(endDate));
        const title = `Attendance Report: ${startDateFormatted} - ${endDateFormatted}`;

        doc.font('Helvetica-Bold').fontSize(14);
        doc.text(title, { align: 'center' });
        doc.moveDown(0.5);

        const dateRange = eachDayOfInterval({ start: new Date(startDate), end: endOfDay }).map(date => format(date, 'dd'));

        const facultyMap = {};

        attendanceData.forEach(record => {
            const { date, faculty, attendance } = record;
            const formattedDate = format(date, 'dd');

            if (!facultyMap[faculty]) {
                facultyMap[faculty] = {
                    students: {}
                };
            }

            for (const studentId in attendance) {
                const studentData = attendance[studentId];
                if (!facultyMap[faculty].students[studentData.name]) {
                    facultyMap[faculty].students[studentData.name] = {};
                }
                facultyMap[faculty].students[studentData.name][formattedDate] = studentData.status;
            }
        });

        let y = 120;
        const leftMargin = 35;
        const lMargin = 110;
        const columnWidth = 16;

        doc.font('Helvetica').fontSize(10);

        Object.keys(facultyMap).forEach((faculty, index) => {
            const { students } = facultyMap[faculty];

            y += 10;
            doc.font('Helvetica-Bold').fontSize(12).text(`Faculty: ${faculty}`, leftMargin, y, { width: 650, text: 'center' });
            y += 20;

            doc.lineWidth(0.5);
            doc.rect(leftMargin, y, dateRange.length * columnWidth, 20).stroke();
            doc.font('Helvetica-Bold').fontSize(10).text('Student', leftMargin + 2, y + 3, { width: lMargin - leftMargin - 4, text: 'center' });
            dateRange.forEach((date, dayIndex) => {
                const x = lMargin + dayIndex * columnWidth;
                doc.font('Helvetica-Bold').fontSize(10).text(date, x, y + 3, { width: columnWidth, text: 'center' });
                doc.rect(x, y, columnWidth, 20).stroke();
            });

            y += 20;

            Object.keys(students).forEach((studentName, studentIndex) => {
                doc.font('Helvetica').fontSize(10).text(studentName, leftMargin, y);
                doc.rect(leftMargin, y, lMargin - leftMargin + dateRange.length * columnWidth, 15).stroke();

                dateRange.forEach((date, dayIndex) => {
                    const formattedDate = `${date}`;
                    const status = students[studentName][formattedDate] || 'A';
                    const color = status === 'P' ? 'green' : 'red';
                    const x = lMargin + dayIndex * columnWidth + (columnWidth / 2) - (doc.widthOfString(status) / 2);
                    doc.fillColor(color).text(status, x, y + 3, { width: columnWidth, align: 'center' });
                    doc.fillColor('black');
                    doc.rect(lMargin + dayIndex * columnWidth, y, columnWidth, 15).stroke();
                });

                y += 15;
            });

            y += 10;
        });

        doc.end();
    } catch (error) {
        console.error("Error exporting attendance as PDF:", error);
        res.status(500).send("Internal Server Error");
    }
});

function formatDate(date) {
    const options = { year: 'numeric', month: 'long', day: 'numeric' };
    return date.toLocaleDateString('en-GB', options);
}

// exportAttendanceJSON
app.get("/exportAttendanceJSON/:startDate/:endDate", async (req, res) => {
    try {
        const { startDate, endDate } = req.params;
        const { faculties } = req.query; 
        const endOfDay = new Date(endDate);
        endOfDay.setHours(23, 59, 59, 999);
        let filter = {
            date: {
                $gte: new Date(startDate),
                $lte: endOfDay
            }
        };
        if (faculties) {
            filter.faculty = { $in: faculties.split(",") };
        }
        const attendanceData = await attendanceModel.find(filter);
        const jsonAttendanceData = attendanceData.map(record => {
            const { date, faculty, attendance } = record;
            const formattedDate = format(date, 'dd-MM-yyyy');
            const formattedAttendance = {};
            for (const studentId in attendance) {
                const studentData = attendance[studentId];
                formattedAttendance[studentData.name] = {
                    date: formattedDate,
                    status: studentData.status
                };
            }
            return {
                faculty,
                attendance: formattedAttendance
            };
        });
        res.json(jsonAttendanceData);
    } catch (error) {
        console.error("Error exporting attendance:", error);
        res.status(500).send("Internal Server Error");
    }
});


