import { isDefaultClause, transpileModule, unescapeLeadingUnderscores } from 'typescript';
import * as XLSX from 'xlsx';

const kprFile = document.getElementById('kprFile');
const kirFile = document.getElementById('kirFile');
const outputDiv = document.getElementById('output');
const errorDiv = document.getElementById('error');
const downloadJson = document.getElementById('downloadJson');

var kpr = undefined;
var kir = undefined;
var furs_json = undefined;

function parse_simple_value(raw) {
    if (raw == undefined) {
        return {value: undefined, raw: undefined};
    } 

    return {
        value: raw.v,
        raw: raw,
    }
}


function parse_integer(raw) {
    return parse_simple_value(raw);
}

function parse_string(raw) {
    if (raw == undefined) {
        return {value: undefined, raw: undefined};
    } 

    return {
        value: raw.w,
        raw: raw,
    }
}

function parse_date(raw) {
    return parse_simple_value(raw);
}

function parse_float(raw) {
    var data = parse_integer(raw);
    if (data.value == undefined) {
        data.value = 0.0;
    }
    return data;
}

function parse_file(data, start_row) {
    var parsed_data = [];
    const workbook = XLSX.read(data, { type: 'array', cellDates: true });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    var row = start_row;
    while (worksheet['A' + row]) {
        parsed_data.push({
            zaporedna: parse_integer(worksheet['A' + row]),
            datum_knjizenja: parse_date(worksheet['B' + row]),
            stevilka_racuna: parse_integer(worksheet['C' + row]),
            datum_racuna: parse_date(worksheet['D' + row]),
            podjetje: parse_string(worksheet['E' + row]),
            davcna: parse_string(worksheet['F' + row]),
            vrednost_z_ddv: parse_float(worksheet['G' + row]),
            izvoz_usa: parse_float(worksheet['H' + row]),
            dobave_v_eu: parse_float(worksheet['I' + row]),
            druge_oprostitve: parse_float(worksheet['J' + row]),
            osnova_9_5: parse_float(worksheet['K' + row]),
            ddv_9_5: parse_float(worksheet['L' + row]),
            osnova_22: parse_float(worksheet['M' + row]),
            ddv_22: parse_float(worksheet['N' + row]),
        });

        row++;
    }

    parsed_data.sort((a, b) => a.zaporedna - b.zaporedna);
    console.log(parsed_data);
    return parsed_data;
}

function validate_file(event) {
    const file = event.target.files[0];
    if (!file) {
        return false;
    }

    // Validate file type
    const validTypes = [
        'application/vnd.ms-excel',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    ];
    if (!validTypes.includes(file.type)) {
        errorDiv.textContent = 'Omogočeno je samo nalaganje Excel datotek (.xls or .xlsx)';
        outputDiv.innerHTML = '';
        downloadJson.style.display = 'none';
        return false;
    }

    return true;
}

function generate_furs_json() {
    if (kir == undefined || kpr == undefined) {
        return
    }

    furs_json = {
        // TODO put correct data into the metadata header.
        Glava: {
            TaxPayerID: "12345678",
            TUJEC1: "AB",
            TUJEC2: "string",
            OBDOBJE_OD: "2024-06-01T00:00:00.0000001+02:00",
            OBDOBJE_DO: "2024-06-30T00:00:00.0000001+02:00",
            KIR: "true",
            KPR: "true",
            VRACILO: "false",
            ODBDELEZ: "false",
            NACIN: 3,
            INSPOS: "false",
            PREDLODO: "false",
            OPOMBA: "string"
        },
        Lista_KIR: {
            KIR: []
        },
        Lista_KPR: {
            KPR: []
        }
    };

    kir.forEach((row) => {
        furs_json.Lista_KIR.KIR.push({
            ZAPST: row.zaporedna.value,
            OBDOBJE: "0123", // TODO
            P2: row.datum_knjizenja.value.toISOString(),
            P3: row.stevilka_racuna.value,
            P4: row.datum_racuna.value.toISOString(),
            P5: row.podjetje.value,
            P6: "SI", // TODO
            P6DS: row.davcna.value,
            P7: 0.00, // TODO
            P8: 0.00, // TODO
            P9: 0.00, // TODO
            P10: 0.00, // TODO
            P11: 0.00, // TODO
            P12: 0.00, // TODO
            P13: 0.00, // TODO
            P14: 0.00, // TODO
            P15: 0.00, // TODO
            P16: 0.00, // TODO
            P17: 0.00, // TODO
            P18: 0.00, // TODO
            P19: 0.00, // TODO
            P20: 0.00, // TODO
            P21: 0.00, // TODO
            P22: 0.00, // TODO
            P23: 0.00, // TODO
            P24: 0.00, // TODO
            P25: 0.00, // TODO
            P26: 0.00, // TODO
            P27: 0.00, // TODO
            P29: "Opombe", // TODO
            OBRAVNAVA: 1, // TODO
            OBDOBJE88: "01022024", // TODO
            DAVEK88: 0.00, // TODO
        });
    })


    kpr.forEach((row) => {
        furs_json.Lista_KPR.KPR.push({
            ZAPST: row.zaporedna.value,
            OBDOBJE: "0123", // TODO
            P2: row.datum_knjizenja.value.toISOString(),
            P3: row.stevilka_racuna.value,
            P4: "2024-06-01T00:00:00.0000001+02:00", // TODO
            P5: row.datum_racuna.value.toISOString(),
            P6: row.podjetje.value,
            P7: "SI", // TODO
            P7DS: row.davcna.value,
            P8: 0.00, // TODO
            P9: 0.00, // TODO
            P10: 0.00, // TODO
            P11: 0.00, // TODO
            P12: 0.00, // TODO
            P13: 0.00, // TODO
            P14: 0.00, // TODO
            P15: 0.00, // TODO
            P16: 0.00, // TODO
            P17: 0.00, // TODO
            P18: 0.00, // TODO
            P19: 0.00, // TODO
            P20: 0.00, // TODO
            P21: 0.00, // TODO
            P22: "Opombe", // TODO
            OBRAVNAVA: 1, // TODO
            OBDOBJE88: "01022024", // TODO
            DAVEK88: 0.00, // TODO
        });
    })

    // Prepare JSON download
    const jsonString = JSON.stringify(furs_json, null, 2);
    const blob = new Blob([jsonString], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    downloadJson.href = url;
    downloadJson.style.display = 'inline-block';
    // TODO name file using company name and period.

    // Clean up URL after download
    // downloadJson.addEventListener('click', () => {
    //     setTimeout(() => URL.revokeObjectURL(url), 100);
    // }, { once: true });
}

kprFile.addEventListener('change', async (event) => {
    if (!validate_file(event)) {
        return;
    }

    try {
        // Read and parse the file
        kpr = parse_file(await event.target.files[0].arrayBuffer(), 10);
        generate_furs_json();
    } catch (error) {
        errorDiv.textContent = 'Error parsing file: ' + error.message;
        outputDiv.innerHTML = '';
        downloadJson.style.display = 'none';
    }
});


kirFile.addEventListener('change', async (event) => {
    if (!validate_file(event)) {
        return;
    }

    try {
        // Read and parse the file
        kir = parse_file(await event.target.files[0].arrayBuffer(), 9);
        generate_furs_json();



        // Clear précédente output and errors
        // errorDiv.textContent = '';
        // outputDiv.innerHTML = '';

        // Create table
        // const table = document.createElement('table');
        // jsonData.forEach((row, index) => {
        //     const tr = document.createElement('tr');
        //     row.forEach(cell => {
        //         const cellElement = index === 0 ? 'th' : 'td';
        //         const td = document.createElement(cellElement);
        //         td.textContent = cell ?? '';
        //         tr.appendChild(td);
        //     });
        //     table.appendChild(tr);
        // });

        // outputDiv.appendChild(table);

    } catch (error) {
        errorDiv.textContent = 'Error parsing file: ' + error.message;
        outputDiv.innerHTML = '';
        downloadJson.style.display = 'none';
    }
});
