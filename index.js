import express from "express";
import cors from "cors";
import mongoose from "mongoose";
import "dotenv/config";
import ExcelJS from "exceljs";

import { format, eachDayOfInterval } from 'date-fns';
const port = 3000;
// import { format } from 'date-fns';
// import { toZonedTime } from 'date-fns-tz'
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

// ........................... 


//   ............................ 

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
        const result = await facultyModel.deleteOne({
            "facultyList.id": facultyIdToDelete
        });

        if (result) {
            console.log("Faculty deleted", result);
            res.json("Faculty Deleted");
        }
    } catch (error) { console.log(error) }
});

app.delete("/deleteFacultyMany", async (req, res) => {
    const facultyIdsToDelete = req.body.ids;
    // console.log(facultyIdsToDelete)
    try {
        const result = await facultyModel.deleteMany({
            "facultyList.id": { $in: facultyIdsToDelete }
        })
        if (result) console.log(result)

    } catch (error) { console.log(error) }

});


// try {
//     const result = await facultyModel.updateMany(
//         { "facultyList.id": { $in: facultyIdsToDelete } },
//         { $pull: { facultyList: { id: { $in: facultyIdsToDelete } } } }
//     );

//     if (result.modifiedCount > 0) {
//         res.json("Faculty Deleted");
//     } else {
//         res.json(false);
//     }
// } catch (error) {
//     console.log("Error deleting faculty:", error);
//     res.json(false);
// }


// const result = await facultyModel.updateOne(
//     { "facultyList.id": facultyIdToDelete },
//     { $pull: { facultyList: { id: facultyIdToDelete } } }
// );
// if (result.modifiedCount > 0) {
//     res.json("Faculty Deleted");
// } else {
//     res.json(false);

// }

app.get("/exportAttendance/:startDate/:endDate", async (req, res) => {
    try {
        const { startDate, endDate } = req.params;

        const endOfDay = new Date(endDate);
        endOfDay.setHours(23, 59, 59, 999);

        const attendanceData = await attendanceModel.find({
            date: {
                $gte: new Date(startDate),
                $lte: endOfDay
            }
        });

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet("Attendance");

        const formattedStartDate = format(new Date(startDate), 'dd-MM-yyyy', { timeZone: 'Asia/Kolkata' });
        const formattedEndDate = format(new Date(endDate), 'dd-MM-yyyy', { timeZone: 'Asia/Kolkata' });
        worksheet.mergeCells('A1:C1');
        worksheet.getCell('A1').value = `Attendance from ${formattedStartDate} to ${formattedEndDate}`;
        worksheet.getCell('A1').alignment = { horizontal: 'center' };

        const dateRange = eachDayOfInterval({ start: new Date(startDate), end: new Date(endOfDay) }).map(date => format(date, 'dd-MM-yyyy'));

        worksheet.addRow(["Faculty", "Student Name", "Date", ...dateRange]);

        worksheet.getColumn(1).width = 20;
        worksheet.getColumn(2).width = 20;
        worksheet.getColumn(3).width = 15;
        for (let i = 4; i < dateRange.length + 4; i++) {
            worksheet.getColumn(i).width = 15;
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

                dateRange.forEach((date, index) => {
                    const cell = newRow.getCell(index + 4);
                    if (cell.value === 'P') {
                        cell.font = { color: { argb: 'FF00FF00' } };
                    } else if (cell.value === 'A') {
                        cell.font = { color: { argb: 'FFFF0000' } };
                    }
                });

                isFirstStudent = false;
            });
        });

        const filename = `attendance_${startDate}_${endDate}.xlsx`;

        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
        await workbook.xlsx.write(res);

        res.end();
    } catch (error) {
        console.error("Error exporting attendance:", error);
        res.status(500).send("Internal Server Error");
    }
});