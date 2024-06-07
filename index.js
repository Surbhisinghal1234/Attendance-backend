import express from "express";
import cors from "cors";
import mongoose from "mongoose";
import "dotenv/config";

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
    name: {
        type: String,
        required: true,
    },
    id: {
        type: Number,
        default: Date.now(),
    },
});

const studentSchema = new mongoose.Schema({
    id: {
        type: Number,
        default: Date.now(),
    },
    name: {
        type: String,
        required: true,
    },
    faculty: {
        type: String,
        required: true,
    },
    aadharCard: { 
        type: String,
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

app.get("/getFaculty", async (req, res) => {
    const facultyList = await facultyModel.find({}).select("-__v");
    if (facultyList) res.json(facultyList);
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

    dataToSave
        .save()
        .then((response) => res.json(response))
        .catch((error) => {
            console.log(error);
            res.send(false);
            console.log("dataToSave", dataToSave);

        });
});

app.delete("/deleteStudent/:id", async (req, res) => {
    const idToDelete = req.params.id;
    const deletedStudent = await studentModel.findOneAndDelete({
        id: idToDelete
    });
    if (deletedStudent) res.json("Student Deleted");
    else res.json(false);
});

app.delete("/deleteFaculty/:id", async (req, res) => {
    const idToDelete = req.params.id;
    try {
        const deletedFaculty = await facultyModel.findOneAndDelete({ _id: idToDelete });
        if (deletedFaculty) {
            res.json("Faculty Deleted");
        } else {
            res.json(false);
        }
    } catch (error) {
        console.log("error", error);
        res.json(false);
    }
});




