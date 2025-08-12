import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { format as formatDate } from "date-fns";

const kprFile = document.getElementById('kprFile');
const kirFile = document.getElementById('kirFile');
const outputDiv = document.getElementById('output');
const errorDiv = document.getElementById('error');
const downloadJson = document.getElementById('downloadJson');
const downloadXml = document.getElementById('downloadXml');
const downloadZip = document.getElementById('downloadZip');
const resetForm = document.getElementById('resetForm');
const metadataDiv = document.getElementById('metadata');
const filesDiv = document.getElementById('files');

const davcnaStevilka = document.getElementById('davcnaStevilka');
const zahtevamVracilo = document.getElementById('zahtevamVracilo');
const izracunavamOdbitniDelez = document.getElementById('izracunavamOdbitniDelez');
const nacin = document.getElementById('nacin');
const opomba = document.getElementById('opomba');
const obdobjeOd = document.getElementById('obdobjeOd');
const obdobjeDo = document.getElementById('obdobjeDo');

var kpr = undefined;
var kir = undefined;

var version = 'development_version';
var versionElement = document.getElementById('appVersion');
versionElement.innerHTML = version;

function generateTable(data, inverted = false) {
    const table = document.createElement('table');
    if (inverted) {
        for (var key in data) {
            const tr = document.createElement('tr');
            var td = document.createElement('th');
            td.textContent = key ?? '';
            tr.appendChild(td);
            table.appendChild(tr);

            var td = document.createElement('td');
            td.textContent = data[key] ?? '';
            tr.appendChild(td);
            table.appendChild(tr);
        }
    } else {
        data.forEach((row, index) => {
            if (index === 0) {
                const tr = document.createElement('tr');
                for (var key in row) {
                    const td = document.createElement('th');
                    td.textContent = key ?? '';
                    tr.appendChild(td);
                }
                table.appendChild(tr);
            }
            const tr = document.createElement('tr');
            for (var key in row) {
                const td = document.createElement('td');
                td.textContent = row[key] ?? '';
                tr.appendChild(td);
            }
            table.appendChild(tr);
        });
    }
    return table;
}

function setDefaultDates() {
    const today = new Date();
    const prevMonthFirstDay = new Date(today.getFullYear(), today.getMonth() - 1, 1);
    const prevMonthLastDay = new Date(today.getFullYear(), today.getMonth(), 0);
    document.getElementById('obdobjeOd').value = formatDate(prevMonthFirstDay, "yyyy-MM-dd");
    document.getElementById('obdobjeDo').value = formatDate(prevMonthLastDay, "yyyy-MM-dd");
}

function parse_simple_value(raw) {
    if (!raw) {
        return {value: undefined, raw: undefined};
    }

    return {
        value: raw.v,
        raw: raw,
    }
}

function parse_integer(raw) {
    if (!raw) {
        return {value: 0, raw: 0};
    }

    var data = parse_simple_value(raw);
    if (!data.value || (typeof data.value === "string" && !data.value.trim())) {
        return {value: 0, raw: 0};
    }

    return data;

}

function parse_string(raw) {
    if (!raw) {
        return {value: '', raw: ''};
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
    if (!raw) {
        return {value: 0.0, raw: 0.0};
    }

    var data = parse_simple_value(raw);
    if (!data.value || (typeof data.value === "string" && !data.value.trim())) {
        return {value: 0.0, raw: 0.0};
    }

    return data;
}

function parse_kir(data) {
    var parsed_data = [];
    const workbook = XLSX.read(data, { type: 'array', cellDates: true });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    var row = 9;
    while (worksheet['A' + row]) {
        parsed_data.push({
            zaporedna: parse_integer(worksheet['A' + row]),             // ZAPST
            datum_knjizenja: parse_date(worksheet['B' + row]),          // P2
            stevilka_racuna: parse_integer(worksheet['C' + row]),       // P3
            datum_racuna: parse_date(worksheet['D' + row]),             // P4
            podjetje: parse_string(worksheet['E' + row]),               // P5
            koda_drzave: parse_string(worksheet['F' + row]),            // P6
            davcna: parse_string(worksheet['G' + row]),                 // P6DS
            vrednost_brez_ddv: parse_float(worksheet['H' + row]),       // P7
            ddv_prejemnik: parse_float(worksheet['I' + row]),           // P8
            oproscene_dobave_slo: parse_float(worksheet['J' + row]),    // P9
            oproscene_dobave_eu: parse_float(worksheet['K' + row]),     // P10
            tristranske_dobave_eu: parse_float(worksheet['L' + row]),   // P11
            prodaja_na_daljavo: parse_float(worksheet['M' + row]),      // P12
            dobava_v_eu: parse_float(worksheet['N' + row]),             // P13
            ddv_22: parse_float(worksheet['O' + row]),                  // P14
            ddv_9: parse_float(worksheet['P' + row]),                   // P15
            ddv_5: parse_float(worksheet['Q' + row]),                   // P16
            prid_material_eu_22: parse_float(worksheet['R' + row]),     // P17
            prid_storitve_eu_22: parse_float(worksheet['S' + row]),     // P18
            prid_material_eu_9: parse_float(worksheet['T' + row]),      // P19
            prid_storitve_eu_9: parse_float(worksheet['U' + row]),      // P20
            prid_material_eu_5: parse_float(worksheet['V' + row]),      // P21
            prid_storitve_eu_5: parse_float(worksheet['W' + row]),      // P22
            samoob_22: parse_float(worksheet['X' + row]),               // P23
            samoob_9: parse_float(worksheet['Y' + row]),                // P24
            samoob_5: parse_float(worksheet['Z' + row]),                // P25
            samoob_uvoz: parse_float(worksheet['AA' + row]),            // P26
            dobave_zunaj_slo: parse_float(worksheet['AB' + row]),       // P27
            opombe: parse_string(worksheet['AC' + row]),                // P28
            obdobje: parse_string(worksheet['N1']),                     // OBDOBJE
            nacin: parse_string(worksheet['S1']),                       // OBRAVNAVA
        });

        row++;
    }

    parsed_data.sort((a, b) => a.zaporedna - b.zaporedna);
    console.log(parsed_data);
    return parsed_data;
}

function parse_kpr(data) {
    var parsed_data = [];
    const workbook = XLSX.read(data, { type: 'array', cellDates: true });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    var row = 10;
    while (worksheet['A' + row]) {
        parsed_data.push({
            zaporedna: parse_integer(worksheet['A' + row]),             // ZAPST
            datum_knjizenja: parse_date(worksheet['B' + row]),          // P2
            stevilka_racuna: parse_integer(worksheet['C' + row]),       // P3
            datum_prejema: parse_date(worksheet['D' + row]),            // P4
            datum_racuna: parse_date(worksheet['E' + row]),             // P5
            podjetje: parse_string(worksheet['F' + row]),               // P6
            koda_drzave: parse_string(worksheet['G' + row]),            // P7
            davcna: parse_string(worksheet['H' + row]),                 // P7DS
            vrednost_brez_ddv: parse_float(worksheet['I' + row]),       // P8
            ddv_obrac_prejem: parse_float(worksheet['J' + row]),        // P9
            pridobitve_blaga_eu: parse_float(worksheet['K' + row]),     // P10
            pridobitve_storitev_eu: parse_float(worksheet['L' + row]),  // P11
            nepremicnine: parse_float(worksheet['M' + row]),            // P12
            osnovna_sredstva: parse_float(worksheet['N' + row]),        // P13
            oproscene_nabave: parse_float(worksheet['O' + row]),        // P14
            oproscene_neprem: parse_float(worksheet['P' + row]),        // P15
            oproscena_oprema: parse_float(worksheet['Q' + row]),        // P16
            ne_obije: parse_float(worksheet['R' + row]),                // P17
            ddv_22: parse_float(worksheet['S' + row]),                  // P18
            ddv_9: parse_float(worksheet['T' + row]),                   // P19
            ddv_5: parse_float(worksheet['U' + row]),                   // P20
            pavsal_8: parse_float(worksheet['V' + row]),                // P21
            opombe: parse_string(worksheet['W' + row]),                 // P22
            obdobje: parse_string(worksheet['B1']),                     // OBDOBJE
            nacin: parse_string(worksheet['D1']),                       // OBRAVNAVA
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
    // Note: This does not pass validation on eDavki and I can't figure out why. Keeping for
    // historical reasons. We are using XML instead.
    var furs_json = {
        Glava: {
            TaxPayerID: davcnaStevilka.value.trim(),
            // TODO do we need this?
            //TUJEC1: "AB",
            //TUJEC2: "string",
            // TODO How can we only get this once?
            OBDOBJE_OD: formatDate(obdobjeOd.valueAsDate, "yyyy-MM-dd"),
            OBDOBJE_DO: formatDate(obdobjeDo.valueAsDate, "yyyy-MM-dd"),
            // TODO se zgodi da je kateri od KIR/KPR prazen?
            KIR: true,
            KPR: true,
            VRACILO: zahtevamVracilo.checked,
            ODBDELEZ: izracunavamOdbitniDelez.checked,
            INSPOS: false,
            PREDLODO: false,
            OPOMBA: opomba.value,
        },
        Lista_KIR: {
            KIR: []
        },
        Lista_KPR: {
            KPR: []
        }
    };

    if (nacin.value.trim() != '') {
        furs_json.Glava.NACIN = nacin.value.trim();
    }

    kir.forEach((row) => {
        furs_json.Lista_KIR.KIR.push({
            ZAPST: row.zaporedna.value,
            OBDOBJE: row.obdobje.value,
            P2: formatDate(row.datum_knjizenja.value, "yyyy-MM-dd"),
            P3: row.stevilka_racuna.value,
            P4: formatDate(row.datum_racuna.value, "yyyy-MM-dd"),
            P5: row.podjetje.value,
            P6: row.koda_drzave.value,
            P6DS: row.davcna.value,
            P7: row.vrednost_brez_ddv.value,
            P8: row.ddv_prejemnik.value,
            P9: row.oproscene_dobave_slo.value,
            P10: row.oproscene_dobave_eu.value,
            P11: row.tristranske_dobave_eu.value,
            P12: row.prodaja_na_daljavo.value,
            P13: row.dobava_v_eu.value,
            P14: row.ddv_22.value,
            P15: row.ddv_9.value,
            P16: row.ddv_5.value,
            P17: row.prid_material_eu_22.value,
            P18: row.prid_storitve_eu_22.value,
            P19: row.prid_material_eu_9.value,
            P20: row.prid_storitve_eu_9.value,
            P21: row.prid_material_eu_5.value,
            P22: row.prid_storitve_eu_5.value,
            P23: row.samoob_22.value,
            P24: row.samoob_9.value,
            P25: row.samoob_5.value,
            P26: row.samoob_uvoz.value,
            P27: row.dobave_zunaj_slo.value,
            P28: row.opombe.value,
            OBRAVNAVA: row.nacin.value,
        });
    })

    kpr.forEach((row) => {
        furs_json.Lista_KPR.KPR.push({
            ZAPST: row.zaporedna.value,
            OBDOBJE: row.obdobje.value,
            P2: formatDate(row.datum_knjizenja.value, "yyyy-MM-dd"),
            P3: row.stevilka_racuna.value,
            P4: formatDate(row.datum_prejema.value, "yyyy-MM-dd"),
            P5: formatDate(row.datum_racuna.value, "yyyy-MM-dd"),
            P6: row.podjetje.value,
            P7: row.koda_drzave.value,
            P7DS: row.davcna.value,
            P8: row.vrednost_brez_ddv.value,
            P9: row.ddv_obrac_prejem.value,
            P10: row.pridobitve_blaga_eu.value,
            P11: row.pridobitve_storitev_eu.value,
            P12: row.nepremicnine.value,
            P13: row.osnovna_sredstva.value,
            P14: row.oproscene_nabave.value,
            P15: row.oproscene_neprem.value,
            P16: row.oproscena_oprema.value,
            P17: row.ne_obije.value,
            P18: row.ddv_22.value,
            P19: row.ddv_9.value,
            P20: row.ddv_5.value,
            P21: row.pavsal_8.value,
            P22: row.opombe.value,
            OBRAVNAVA: row.nacin.value,
        });
    })

    return furs_json;
}

function generate_furs_xml(export_data) {
    const emptyValuesToOmit = [
        '.DDV_KIR_KPR.Lista_KIR.KIR.P6',
        '.DDV_KIR_KPR.Lista_KIR.KIR.P6DS',
        '.DDV_KIR_KPR.Lista_KIR.KIR.P28',

        '.DDV_KIR_KPR.Lista_KPR.KPR.P7',
        '.DDV_KIR_KPR.Lista_KPR.KPR.P22',
    ];

    function convertToXML(data, parentTag, path = '', indent = '') {
        let xml = '';

        const parentTagOpen = parentTag === "DDV_KIR_KPR" ?
            `DDV_KIR_KPR xmlns="http://edavki.durs.si/Documents/Schemas/DDV_KIR_KPR_1.xsd" xsi:schemaLocation="http://edavki.durs.si/Documents/Schemas/DDV_KIR_KPR_1.xsd schema.xsd" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"` :
            parentTag;

        if (Array.isArray(data)) {
            // Handle arrays
            data.forEach(item => {
                xml += convertToXML(item, parentTag, path, indent);
            });
        } else if (typeof data === 'object' && data !== null) {
            // Handle objects
            xml += `${indent}<${parentTagOpen}>\n`;
            for (const [key, value] of Object.entries(data)) {
                xml += convertToXML(value, key, path + '.' + parentTag, indent + '  ');
            }
            xml += `${indent}</${parentTag}>\n`;
        } else {
            if (
                !emptyValuesToOmit.includes(path + '.' + parentTag) ||
                (data && (typeof data.value === "string" && data.value.trim()))
            ) {
                xml += `${indent}<${parentTagOpen}>${data}</${parentTag}>\n`;
            }
        }

        return xml;
    }

    // Start XML document
    let xmlString = `<?xml version="1.0" encoding="utf-8"?>`;
    xmlString += convertToXML(export_data, "DDV_KIR_KPR");

    return xmlString;
}

function generate_furs_files() {
    // TODO validate tax id?
    if (kir == undefined || kpr == undefined || davcnaStevilka.value.trim() == '' || obdobjeOd.value.trim() == '' || obdobjeDo.value.trim() == '') {
        downloadJson.style.display = 'none';
        downloadXml.style.display = 'none';
        downloadZip.style.display = 'none';
        resetForm.style.display = 'none';
        outputDiv.innerHTML = '';
        return;
    }

    // Prepare JSON download
    const export_data = generate_furs_json();
    const jsonString = JSON.stringify(export_data, null, 2);
    const jsonBlob = new Blob([jsonString], { type: 'application/json' });
    const jsonFilename = `DDV_${davcnaStevilka.value.trim()}_${formatDate(obdobjeOd.value, "yyyyMM")}_${formatDate(obdobjeDo.value, "yyyyMM")}.json`;

    // Prepare XML download.
    const xmlString = generate_furs_xml(export_data);
    const xmlBlob = new Blob([xmlString], { type: 'application/xml' });
    const xmlFilename = `DDV_${davcnaStevilka.value.trim()}_${formatDate(obdobjeOd.value, "yyyyMM")}_${formatDate(obdobjeDo.value, "yyyyMM")}.xml`;

    // Create ZIP file
    const zip = new JSZip();
    zip.file(xmlFilename, xmlString);
    zip.generateAsync({ type: 'blob' }).then(function(zipBlob) {
        downloadZip.href = URL.createObjectURL(zipBlob);
        downloadZip.download = `DDV_${davcnaStevilka.value.trim()}_${formatDate(obdobjeOd.value, "yyyyMM")}_${formatDate(obdobjeDo.value, "yyyyMM")}.zip`;
        downloadZip.style.display = 'inline-block';
        resetForm.style.display = 'inline-block';
    });

    downloadJson.href = URL.createObjectURL(jsonBlob);
    downloadJson.download = jsonFilename;

    downloadXml.href = URL.createObjectURL(xmlBlob);
    downloadXml.download = xmlFilename;

    // Create tables
    outputDiv.innerHTML = '';
    var heading = document.createElement('h3');
    heading.innerHTML = 'Glava';
    outputDiv.appendChild(heading);
    outputDiv.appendChild(generateTable(export_data.Glava, true));
    var heading = document.createElement('h3');
    heading.innerHTML = 'Knjiga izdanih računov';
    outputDiv.appendChild(heading);
    outputDiv.appendChild(generateTable(export_data.Lista_KIR.KIR));
    var heading = document.createElement('h3');
    heading.innerHTML = 'Knjiga prejetih računov';
    outputDiv.appendChild(heading);
    outputDiv.appendChild(generateTable(export_data.Lista_KPR.KPR));

    metadataDiv.style.display = 'none';
    filesDiv.style.display = 'none';
}

kprFile.addEventListener('change', async (event) => {
    if (!validate_file(event)) {
        return;
    }

    try {
        kpr = parse_kpr(await event.target.files[0].arrayBuffer());
        generate_furs_files();
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
        kir = parse_kir(await event.target.files[0].arrayBuffer());
        generate_furs_files();
    } catch (error) {
        errorDiv.textContent = 'Error parsing file: ' + error.message;
        outputDiv.innerHTML = '';
        downloadJson.style.display = 'none';
    }
});

resetForm.addEventListener('click', async (_event) => {
    downloadJson.style.display = 'none';
    downloadXml.style.display = 'none';
    downloadZip.style.display = 'none';
    resetForm.style.display = 'none';
    metadataDiv.style.display = 'block';
    filesDiv.style.display = 'flex';
    outputDiv.innerHTML = '';
    setDefaultDates();
    davcnaStevilka.value = '';
    zahtevamVracilo.checked = false;
    izracunavamOdbitniDelez.checked = false;
    nacin.value = '';
    opomba.value = '';
    kprFile.value = '';
    kirFile.value = '';
    kir = undefined;
    kpr = undefined;
});

davcnaStevilka.addEventListener('change', async (event) => {
    generate_furs_files();
});
zahtevamVracilo.addEventListener('change', async (event) => {
    generate_furs_files();
});
izracunavamOdbitniDelez.addEventListener('change', async (event) => {
    generate_furs_files();
});
nacin.addEventListener('change', async (event) => {
    generate_furs_files();
});
opomba.addEventListener('change', async (event) => {
    generate_furs_files();
});
obdobjeOd.addEventListener('change', async (event) => {
    generate_furs_files();
});
obdobjeDo.addEventListener('change', async (event) => {
    generate_furs_files();
});

// Set default dates for Obdobje od (first day of previous month) and Obdobje do (last day of previous month)
document.addEventListener('DOMContentLoaded', setDefaultDates);
