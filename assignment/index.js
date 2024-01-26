const XLSX = require("xlsx");
const fs = require("fs");
const file_path = "Assignment_file.xlsx";
const workbook = XLSX.readFile(file_path);
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];
const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
// this shift is used to remove first element which is title of all value
jsonData.shift();
let result = "";
//  below function used to convert time in format "hh:mm" to total minute
// I have assumed that Time Cards in format 'hh:mm'
function time_spend(a) {
  let hour = "";
  let minute = "";
  if (a == "") {
    return 0;
  } else {
    let find = false;
    for (let i = 0; i < a.length; i++) {
      if (a[i] == ":") {
        find = true;
      } else if (find == true) {
        minute = minute + a[i];
      } else {
        hour = hour + a[i];
      }
    }
  }
  minute = Number(minute);
  hour = Number(hour);
  return hour * 60 + minute;
}

// this loop is used to store the data from excel sheet to a well defined object obj1 defined below
// and I am using map to store data as key-value pair where key is Position_ID and value is obj1
//  obj1={
//     name:"",
//     time_in:[],
//     time_out:[],
//     start_date:"",
//     end_date:"",
//     position_status:[],
//     time:[],
//     }
let map_data = new Map();
for (let i = 0; i < jsonData.length; i++) {
  let obj = {
    Position_ID: "",
    Position_Status: "",
    Timecard_Hours: "",
    Time: "",
    Time_out: "",
    Pay_Cycle_Start_Date: "",
    Pay_Cycle_End_Date: "",
    Employee_Name: "",
  };
  obj.Position_ID = String(jsonData[i][0]);
  obj.Position_Status = jsonData[i][1];
  obj.Time = jsonData[i][2];
  obj.Time_out = jsonData[i][3];
  obj.Timecard_Hours = time_spend(jsonData[i][4]);
  obj.Pay_Cycle_Start_Date = jsonData[i][5];
  obj.Pay_Cycle_End_Date = jsonData[i][6];
  obj.Employee_Name = jsonData[i][7];
  jsonData[i] = obj;

  const id = obj.Position_ID;
  let obj1 = {
    name: "",
    time_in: [],
    time_out: [],
    start_date: "",
    end_date: "",
    position_status: [],
    time: [],
  };
  if (map_data.has(id) === true) {
    let arr = map_data.get(id);
    arr.time.push(obj.Timecard_Hours);
    arr.position_status.push(obj.Position_Status);
    arr.time_in.push(obj.Time);
    arr.time_out.push(obj.Time_out);
  } else {
    obj1.name = obj.Employee_Name;
    obj1.start_date = obj.Pay_Cycle_Start_Date;
    obj1.end_date = obj.Pay_Cycle_End_Date;
    obj1.time.push(obj.Timecard_Hours);
    obj1.position_status.push(obj.Position_Status);
    obj1.time_in.push(obj.Time);
    obj1.time_out.push(obj.Time_out);
    map_data.set(id, obj1);
  }
}

// work for consecutive seven day
result += "(a) Work for 7 consecutive day\n";
console.log("(a) Work for 7 consecutive day");
map_data.forEach((value, key) => {
  let start = 0;
  for (let i = 0; i < value.position_status.length; i++) {
    if (value.position_status[i] === "Active") {
      start = start + 1;
    } else {
      start = 0;
    }
    if (start == 7) {
      let name_list = value.name.split(", ");
      for (let j = 0; j < name_list.length; j++) {
        let namevalue = name_list[j].trim(); // trim the string and assign it back
        result +=
          "Position of Employee:" +
          key +
          "   Name of Employee:" +
          namevalue +
          "\n";
        console.log(
          "Position of Employee:" + key + "   Name of Employee:" + namevalue
        );
      }
      break;
    }
  }
});

// who have less than 10 hours of time between shifts but greater than 1 hour
result +=
  "(b) who have less than 10 hours of time between shifts but greater than 1 hour\n";
console.log(
  "(b) who have less than 10 hours of time between shifts but greater than 1 hour"
);
map_data.forEach((value, key) => {
  if (value.time_in.length == 1) {
    let name_list = value.name.split(", ");
    for (let j = 0; j < name_list.length; j++) {
      let namevalue = name_list[j].trim(); // trim the string and assign it back
      result +=
        "Position of Employee:" +
        key +
        "   Name of Employee:" +
        namevalue +
        "\n";
      console.log(
        "Position of Employee:" + key + "   Name of Employee:" + namevalue
      );
    }
  } else {
    for (let i = 0; i < value.time_in.length - 1; i++) {
      if (value.time_out[i - 1] == "" || value.time_in[i] == "") {
        continue;
      } else {
        let out = Number(value.time_out[i]);
        let enter = Number(value.time_in[i + 1]);
        const val = Math.round((enter - out) * 24 * 60);

        if (val > 60 && val < 600) {
          let name_list = value.name.split(", ");
          for (let j = 0; j < name_list.length; j++) {
            let namevalue = name_list[j].trim(); // trim the string and assign it back
            result +=
              "Position of Employee:" +
              key +
              "   Name of Employee:" +
              namevalue +
              "\n";
            console.log(
              "Position of Employee:" + key + "   Name of Employee:" + namevalue
            );
          }
          break;
        }
      }
    }
  }
});

// Work for more than 14 hours in a single shift
result += "(c) Work for more than 14 hours in a single shift\n";
console.log("(c) Work for more than 14 hours in a single shift");
map_data.forEach((value, key) => {
  for (let i = 0; i < value.position_status.length; i++) {
    const work_time = Number(value.time[i]);
    if (work_time > 14 * 60) {
      let name_list = value.name.split(","); // split names seprated by comma
      for (let j = 0; j < name_list.length; j++) {
        let namevalue = name_list[j].trim(); // trim the string and assign it back
        result +=
          "Position of Employee:" +
          key +
          "   Name of Employee:" +
          namevalue +
          "\n";
        console.log(
          "Position of Employee:" + key + "   Name of Employee:" + namevalue
        );
      }
    }
  }
});

fs.writeFile("output.txt", result, (err) => {
  if (err) {
    console.error("Error in writing file:");
  } else {
    console.log("Content has been written ");
  }
});
