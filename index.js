let jsonObject = {}
let finalData = new Map()
let headers = ['JOB DATE', 'TRIP NUMBER', 'WAY', 'CLIENT NAME', 'SCHEDULED PU', 'PICKUP', 'PU TIME DIFFERENCE', 'PU LATE', 'DO TIME', 'DO TIME DIFF', 'DO LATE', 'CAR NUMBER', 'DRIVER NAME', 'DRIVETIME', 'APT TIME']
const removeSpaces = (data) => data.replace(' ', '')

const ExcelToJSON = function(input) {
    this.parseExcel = function(file) {
        const reader = new FileReader()

        reader.onload = function(e) {
            let data = e.target.result
            let workbook = XLSX.read(data, {type: "binary", raw: false})
            workbook.SheetNames.forEach(function(sheetName) {
                let XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
                console.log(XL_row_object)
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
        let tripNumber = convertVoucher(row['VOUCHER'], row['WAY'])
        finalData.set(tripNumber, {
            'JOB DATE': row['JOB DATE'],
            'TRIP NUMBER': convertVoucher(row['VOUCHER'], row['WAY']),
            'WAY': row['WAY'],
            'CLIENT NAME': row['NAME'],
            'SCHEDULED PU': '',
            'PICKUP': row['PICKUP'],
            'PU TIME DIFFERENCE': '',
            'PU LATE': '',
            'DO TIME': row['DROPOFF'],
            'DO TIME DIFF': '',
            'DO LATE': '',
            'CAR NUMBER': row['CARNUM'],
            'DRIVER NAME': row['DRNAME'],
            'DRIVETIME': row['DRIVE TIME'],
        })
    })
}

function modivCareDefaultData (data, finalData) {
    data['modiv'].forEach((modiv) => {
        let trip = removeSpaces(modiv['Trip ID']);
        if(finalData.get(trip) && finalData.get(trip)['JOB DATE'] === modiv['Trip Date']) {
            finalData.get(trip)['APT TIME'] = modiv['Appointment Time']
        }
    })
}

function convertVoucher (data, leg) {
    let arr = data.split('-')
    arr.splice(1,1)
    return removeSpaces(`${arr.join('-')}-${leg}`)
}

function splitDate (data) {
    data.forEach((row) => {
        let splitDate = row['JOB DATE'].split(' ')
        row['JOB DATE'] = splitDate[0];
        row['SCHEDULED PU'] = splitDate[1]
    })
}

function legTimeDifference (data, tripTime, driverTime, rowDiff, lateRow) {
    data.forEach((row) => {
        let modivPU = moment(row[tripTime],'HH:mm a')
        let driverPU = moment(row[driverTime],'HH:mm a')
        if(modivPU.format('HH:mm a') === moment('00:00', 'HH:mm a').format('HH:mm a')) {
            row[rowDiff] = 'none'
            row[tripTime] = 'none'
            return
        }
        row[rowDiff] = driverPU.diff(modivPU, 'minutes')
        row[tripTime] = modivPU.format('HH:mm a')
        row[driverTime] = driverPU.format('HH:mm a')
        if(+row[rowDiff] > 15) {
            row[lateRow] = 'late'
        }
    })
}

function handleSubmit (evt) {
    evt.preventDefault();
    defaultData(jsonObject)
    splitDate(finalData);
    modivCareDefaultData(jsonObject, finalData)
    legTimeDifference(finalData, "SCHEDULED PU", "PICKUP", "PU TIME DIFFERENCE", 'PU LATE')
    legTimeDifference(finalData, "APT TIME", "DO TIME", "DO TIME DIFF", 'DO LATE')
    console.log(finalData)
}

function handleDownload (evt, data) {
    evt.preventDefault();


}
const altaInput = document.getElementById('alta')
altaInput.addEventListener('change',  (evt) => handleFileSelect(evt,'alta'), false)
const modivInput = document.getElementById('modiv')
modivInput.addEventListener('change',(evt) => handleFileSelect(evt,'modiv'), false)
const buttonSubmit = document.getElementById('submit')
buttonSubmit.addEventListener('click', (evt) => handleSubmit(evt))
const buttonDownload = document.getElementById('download')
buttonDownload.addEventListener('click', (evt) => handleDownload(evt, finalData))