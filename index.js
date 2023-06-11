let jsonObject = {}
let finalData = []

const ExcelToJSON = function(input) {
    this.parseExcel = function(file) {
        const reader = new FileReader()

        reader.onload = function(e) {
            let data = e.target.result
            let workbook = XLSX.read(data, {
                type: 'binary'
            })
            workbook.SheetNames.forEach(function(sheetName) {
                let XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
                let json_object = JSON.stringify(XL_row_object)
                jsonObject[input] = JSON.parse(json_object)
            })
        };
        reader.onerror = function(ex) {
            console.log(ex);
        }
        reader.readAsBinaryString(file);
    }
}

function handleFileSelect(evt, input) {
    const files = evt.target.files;
    const xl2json = new ExcelToJSON(input);
    xl2json.parseExcel(files[0]);
}

function defaultData (data) {
    data['alta'].forEach((row) => {
        finalData.push({
            'JOB DATE': row['JOB DATE'],
            'TRIP NUMBER': convertVoucher(row['VOUCHER'], row['WAY']),
            'WAY': row['WAY'],
            'CLIENT NAME': row['NAME'],
            'SCHEDULED PU A': '',
            'PICKUP': '',
            'A PU TIME DIFFERENCE': '',
            'B LEG DO TIME': row['DROPOFF'],
            'B DO TIME DIFF': '',
            'CAR NUMBER': row['CARNUM'],
            'DRIVER NAME': row['DRNAME'],
            'DRIVETIME': row['DRIVE TIME'],
        })
    })
    data['modiv'].forEach((modiv) => {
        finalData.forEach((alta) => {
            if (modiv['Trip ID'] === alta['TRIP NUMBER'] 
                && modiv['Rider Name'] === alta['CLIENT NAME'] 
                && modiv['Trip Date'] === alta['JOB DATE']) {
                    alta['APT TIME'] = modiv['Appointment Time']
                }
        })
        
    })
}

function convertVoucher (data, leg) {
    let arr = data.split('-')
    arr.splice(1,1)
    return `${arr.join('-')}-${leg}`
}

function splitDate (data) {
    data['alta'].forEach((row) => {
        let splitDate = row['JOB DATE'].split(' ')
        row['JOB DATE'] = splitDate[0];
        row['SCHEDULED PICKUP'] = splitDate[1]
    })
}

function legATimeDifference (data) {
    data['alta'].forEach((row) => {
        let modivPU = moment(row["SCHEDULED PICKUP"],'HH:mm a')
        let driverPU = moment(row["PICKUP"],'HH:mm a')
        row["A PU TIME DIFFERENCE"] = driverPU.diff(modivPU, 'minutes')
        row['PICKUP'] = modivPU.format('HH:mm a')
        row['SCHEDULED PICKUP'] = driverPU.format('HH:mm a')
    })
    console.log(data)
}
function handleSubmit (evt) {
    evt.preventDefault();
    splitDate(jsonObject);
    legATimeDifference(jsonObject)
    defaultData(jsonObject)
    console.log(finalData)
}
const altaInput = document.getElementById('alta')
altaInput.addEventListener('change',  (evt) => handleFileSelect(evt,'alta'), false)
const modivInput = document.getElementById('modiv')
modivInput.addEventListener('change',(evt) => handleFileSelect(evt,'modiv'), false)
const buttonSubmit = document.getElementById('submit')
buttonSubmit.addEventListener('click', (evt) => handleSubmit(evt))