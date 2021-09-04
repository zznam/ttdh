// Requiring the module
const reader = require('xlsx');
const excel = require('excel4node');
const fs = require("fs"); // Or `import fs from "fs";` with ESM

// Reading our test file
const fileName = "20210904-XuLy";

// result file name
const retFileName = fileName + '-Done.xlsx';
var workbook = new excel.Workbook();
var worksheet = workbook.addWorksheet('Sheet 1');
workbook.write(retFileName)

const file = reader.readFile('./' + fileName + '.xlsx')
const mapDistrictBdFile = reader.readFile('./1_mapDistrictBdFile.xlsx')
const provinceCodeFile = reader.readFile('./1_mapProvinceVnFile.xlsx')

let data = []
let mapDistrict = []
let provinceList = []

const sheets = file.SheetNames
const sheetsFile2 = mapDistrictBdFile.SheetNames
const sheetsFile3 = provinceCodeFile.SheetNames

changeAlias = function (str) {
    if (str == null || str == "") return str;
    str = str + "";
    str = str.replace(/à|á|ạ|ả|ã|â|ầ|ấ|ậ|ẩ|ẫ|ă|ằ|ắ|ặ|ẳ|ẵ/g, "a");
    str = str.replace(/è|é|ẹ|ẻ|ẽ|ê|ề|ế|ệ|ể|ễ/g, "e");
    str = str.replace(/Ï|ì|í|ị|ỉ|ĩ|¡/g, "i");
    str = str.replace(/º|ò|ó|ọ|ỏ|õ|ô|ồ|ố|ộ|ổ|ỗ|ơ|ờ|ớ|ợ|ở|ỡ|°/g, "o");
    str = str.replace(/µ|ù|ú|ụ|ủ|ũ|ư|ừ|ứ|ự|ử|ữ/g, "u");
    str = str.replace(/ỳ|ý|ỵ|ỷ|ỹ|¥/g, "y");
    str = str.replace(/đ/g, "d");
    str = str.replace(/À|Á|Ạ|Ả|Ã|Â|Ầ|Ấ|Ậ|Ẩ|Ẫ|Ă|Ằ|Ắ|Ặ|Ẳ|Ẵ/g, "A");
    str = str.replace(/È|É|Ẹ|Ẻ|Ẽ|Ê|Ề|Ế|Ệ|Ể|Ễ/g, "E");
    str = str.replace(/Î|Ì|Í|Ị|Ỉ|Ĩ/g, "I");
    str = str.replace(/Ò|Ó|Ọ|Ỏ|Õ|Ô|Ồ|Ố|Ộ|Ổ|Ỗ|Ơ|Ờ|Ớ|Ợ|Ở|Ỡ/g, "O");
    str = str.replace(/Ù|Ú|Ụ|Ủ|Ũ|Ư|Ừ|Ứ|Ự|Ử|Ữ/g, "U");
    str = str.replace(/Ỳ|Ý|Ỵ|Ỷ|Ỹ/g, "Y");
    str = str.replace(/Đ/g, "D");
    str = str.replace(/€/g, "E");
    str = str.replace(/¬/g, "-");
    str = str.replace(/¹|¼|½/g, "1");
    str = str.replace(/¿|¶|±|±/g, "?");
    str = str.replace(/°/g, "?");
    let ret = "";
    var check1 = /[a-z]/i;
    var check2 = /[A-Z]/i;
    let excluded = ".~` 1234567890-=!@#$%^&*()_+{}[];:<>/";
    for (let i = 0; i < str.length; i++) {
        let e = str.charAt(i);
        if (excluded.indexOf(e) < 0 && !e.match(check1) && !e.match(check2) && /^\d+$/.test(e) == false) {
            e = "?";
        }
        ret += e;
    }
    return ret.toLowerCase().trim();
};

function replaceAll(str, find, replace) {
    const pieces = str.split(find);
    return pieces.join(replace)
}

function phoneNomalize(phone) {
    if (phone != undefined) {
        if (phone.length > 10 && phone.indexOf("/") >= 0)
            phone = phone.split('/')[0];
        phone = phone.trim();
        phone = phone.replace(/\D/g, '');
        phone = Number(phone);
        phone = phone.toString();
        if (phone.length > 9) {
            //Viettel
            if (phone.indexOf("16") == 0) phone = phone.replace("16", "3");
            // MobiFone
            if (phone.indexOf("120") == 0) phone = phone.replace("16", "7");
            if (phone.indexOf("121") == 0) phone = phone.replace("16", "7");
            if (phone.indexOf("122") == 0) phone = phone.replace("16", "7");
            if (phone.indexOf("126") == 0) phone = phone.replace("16", "7");
            if (phone.indexOf("128") == 0) phone = phone.replace("16", "7");
            //VinaPhone
            if (phone.indexOf("123") == 0) phone = phone.replace("16", "8");
            if (phone.indexOf("124") == 0) phone = phone.replace("16", "8");
            if (phone.indexOf("125") == 0) phone = phone.replace("16", "8");
            if (phone.indexOf("127") == 0) phone = phone.replace("16", "8");
            if (phone.indexOf("129") == 0) phone = phone.replace("16", "8");
            //Vietnamobile
            if (phone.indexOf("186") == 0) phone = phone.replace("18", "5");
            if (phone.indexOf("188") == 0) phone = phone.replace("18", "5");
            //Vietnamobile
            if (phone.indexOf("199") == 0) phone = phone.replace("19", "5");
        }

    }
    if (phone == 0) return "";
    return phone
}


for (let i = 0; i < sheets.length; i++) {
    const temp = reader.utils.sheet_to_json(
        file.Sheets[file.SheetNames[i]])
    temp.forEach((res) => {
        data.push(res)
    })
}
for (let i = 0; i < sheetsFile2.length; i++) {
    const temp = reader.utils.sheet_to_json(
        mapDistrictBdFile.Sheets[mapDistrictBdFile.SheetNames[i]])
    temp.forEach((res) => {
        mapDistrict.push(res)
    })
}
for (let i = 0; i < sheetsFile3.length; i++) {
    const temp = reader.utils.sheet_to_json(
        provinceCodeFile.Sheets[provinceCodeFile.SheetNames[i]])
    temp.forEach((res) => {
        provinceList.push(res)
    })
}

// Printing data
// for (let i = 0; i < data.length; i++) {
//     const element = data[i];
//     console.log(element);
// }
console.log("data__length", data.length);
console.log("map__District", mapDistrict.length);
console.log("province__List", provinceList.length);
// for (let i = 0; i < constant.length; i++) {
//     const element = constant[i];
//     console.log(element.namedistrict2);
//     console.log(element.districtCode);
// }

for (let i = 0; i < data.length; i++) {
    const person = data[i];

    let village = person.village ? person.village + ", " : "";
    let commune = person.commune ? person.commune + ", " : "";
    let district = person.district ? person.district : "";
    person.address = village + commune + district;
    person.testCode = (person.testCode).toUpperCase();
    console.log(person.address);

    person.provinceCode = 74;
    let tempD = changeAlias(district).toLowerCase();
    for (let i = 0; i < provinceList.length; i++) {
        const e = provinceList[i];
        let b = changeAlias(e.shortName).toLowerCase();
        if (tempD.indexOf(b) >= 0) {
            person.provinceCode = e.code;
        }
        if (tempD == "tinh ngoai") {
            let a = changeAlias(person.address).toLowerCase();
            if (a.indexOf(b) >= 0) {
                person.provinceCode = e.code;
            }
        }
        if (tempD.indexOf("go vap") >= 0 || tempD.indexOf("hoc mon") >= 0 || tempD.indexOf("cu chi") >= 0 ||
            tempD.indexOf("binh thanh") >= 0 || tempD.indexOf("thu duc") >= 0 ||
            tempD.indexOf("phu nhuan") >= 0 || tempD.indexOf("quan 8") >= 0 ||
            tempD.indexOf("hooc mon") >= 0 || tempD.indexOf("quan 12") >= 0 ||
            tempD.indexOf("quan 4") >= 0 || tempD.indexOf("quan 5") >= 0 ||
            tempD.indexOf("hcm") >= 0) {
            person.provinceCode = 79;
        }
    }
    let phone = person.phone;
    person.phone = phoneNomalize(phone);

    let districtAlias = changeAlias(person.district);
    console.log("district_Alias", districtAlias, i)

    for (let i = 0; i < mapDistrict.length; i++) {
        if (districtAlias == null || districtAlias == "" || districtAlias == " ") {
            person.districtCode = 7202;
            break;
        }
        if (districtAlias.indexOf(changeAlias("bac tan uyen")) >= 0) {
            person.districtCode = 726;
            break;
        }
        if (districtAlias.indexOf(changeAlias("dat cuoc")) >= 0) {
            person.districtCode = 726;
            break;
        }
        if (districtAlias.indexOf(changeAlias("khanh binh")) >= 0) {
            person.districtCode = 723;
            break;
        }

        if (districtAlias.indexOf(changeAlias("khanh loc")) >= 0) {
            person.districtCode = 723;
            break;
        }
        if (districtAlias.indexOf(changeAlias("thoi hoa")) >= 0) {
            person.districtCode = 721;
            break;
        }
        if (districtAlias.indexOf(changeAlias("hoa loi")) >= 0) {
            person.districtCode = 721;
            break;
        }
        if (districtAlias.indexOf(changeAlias("my phuoc")) >= 0 || districtAlias.indexOf(changeAlias("tan dinh")) >= 0) {
            person.districtCode = 721;
            break;
        }
        if (districtAlias.indexOf(changeAlias("mp1")) >= 0 || districtAlias.indexOf(changeAlias("mp2")) >= 0 || districtAlias.indexOf(changeAlias("an tay")) >= 0) {
            person.districtCode = 721;
            break;
        }
        if (districtAlias.indexOf(changeAlias("mp3")) >= 0 || districtAlias.indexOf(changeAlias("mp4")) >= 0 || districtAlias.indexOf(changeAlias("chanh phu hoa")) >= 0) {
            person.districtCode = 721;
            break;
        }
        if (districtAlias.indexOf(changeAlias("tdm")) >= 0) {
            person.districtCode = 718;
            break;
        }
        const con = mapDistrict[i];
        // console.log("nameDistrict1", con.nameDistrict1)
        if (districtAlias.indexOf(changeAlias(con.nameDistrict1)) >= 0) {
            person.districtCode = con.districtCode;
            break;
        }

    }
    if (person.districtCode == "" || person.districtCode == null || person.districtCode == 720 || person.districtCode == 722) {
        person.districtCode = 7202;
    }
    let communeAlias = changeAlias(person.commune);
    console.log("commune_Alias", communeAlias, i);

    for (let i = 0; i < mapDistrict.length; i++) {
        const con = mapDistrict[i];
        // console.log("nameCommune1", con.nameCommune1)
        if (communeAlias == null || communeAlias == "" || communeAlias == " ") {
            person.wardCode = "";
            break;
        }
        if (communeAlias.indexOf(changeAlias("Chanh Phu Hoa")) >= 0) {
            person.wardCode = 25837;
            break;
        }
        if (communeAlias.indexOf(changeAlias("Tan Lap")) >= 0) {
            person.wardCode = 25903;
            break;
        }
        if (communeAlias.indexOf(changeAlias("Tan Long")) >= 0) {
            person.wardCode = 25879;
            break;
        }
        if (districtAlias != null && districtAlias.indexOf(changeAlias("Bac Tan Uyen")) >= 0) {
            if (communeAlias.indexOf(changeAlias("Tan Dinh")) >= 0) {
                person.wardCode = 25894;
                break;
            }
            if (communeAlias.indexOf(changeAlias("Tan Binh")) >= 0) {
                person.wardCode = 25900;
                break;
            }
            if (communeAlias.indexOf(changeAlias("Tan Thanh")) >= 0) {
                person.wardCode = 25906;
                break;
            }
        }
        if (districtAlias != null && districtAlias.indexOf(changeAlias("Di An")) >= 0) {
            if (communeAlias.indexOf(changeAlias("Tan Binh")) >= 0) {
                person.wardCode = 25900;
                break;
            }
            if (communeAlias.indexOf(changeAlias("An Binh")) >= 0) {
                person.wardCode = 25960;
                break;
            }
        }

        if (communeAlias.indexOf(changeAlias(con.nameCommune1)) >= 0) {
            person.wardCode = con.communeCode;
            break;
        }
    }
    if (person.wardCode == "" || person.wardCode == null) {
        person.wardCode = person.provinceCode + "" + person.districtCode;
    }
}

function saveFile(data) {
    let path = './' + retFileName
    console.log("saving_file...");
    if (fs.existsSync(path)) {

        const ws = reader.utils.json_to_sheet(data)
        const fileResult = reader.readFile(path)
        console.log("data,", data.length);
        reader.utils.book_append_sheet(fileResult, ws, "result")
        // Writing to our file
        reader.writeFile(fileResult, path)
        clearInterval(iid)
    }
}
let iid = setInterval(() => {
    saveFile(data)
}, 100);