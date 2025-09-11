import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { format as formatDate } from "date-fns";

const excelFile = document.getElementById('excelFile');
const outputDiv = document.getElementById('output');
const errorDiv = document.getElementById('error');
const errorText = document.getElementById('errorText');
const downloadJson = document.getElementById('downloadJson');
const downloadXml = document.getElementById('downloadXml');
const downloadZip = document.getElementById('downloadZip');
const resetForm = document.getElementById('resetForm');
const fileUploadArea = document.getElementById('fileUploadArea');
const fileName = document.getElementById('fileName');

var version = 'development_version';
var versionElement = document.getElementById('appVersion');
versionElement.innerHTML = version;

function generateTable(data, inverted = false) {
    const wrapper = document.createElement('div');
    wrapper.className = 'table-responsive';

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
                const thead = document.createElement('thead');
                const tr = document.createElement('tr');
                for (var key in row) {
                    const td = document.createElement('th');
                    td.textContent = key ?? '';
                    tr.appendChild(td);
                }
                thead.appendChild(tr);
                table.appendChild(thead);
            }
        });

        const tbody = document.createElement('tbody');
        data.forEach((row, index) => {
            const tr = document.createElement('tr');
            for (var key in row) {
                const td = document.createElement('td');
                td.textContent = row[key] ?? '';
                tr.appendChild(td);
            }
            tbody.appendChild(tr);
        });
        table.appendChild(tbody);
    }

    wrapper.appendChild(table);
    return wrapper;
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

    return {
        value: data.value.toFixed(2),
        raw: data.raw,
    }
}

function parse_boolean(raw) {
    var data = parse_simple_value(raw);
    if (!data.value) {
        return {value: false, raw: false};
    }

    return {
        value: data.value === 'DA',
        raw: data.value === 'DA',
    }
}

function parse_method(raw) {
    var data = parse_simple_value(raw);
    if (!data.value) {
        return {value: null, raw: null};
    }

    const matches = data.value.match(/(?<method>\d) - /);
    if (matches) {
        return {
            value: matches.groups.method,
            raw: matches.groups.method,
        }
    }

    return {value: null, raw: null}
}

function parse_header(data) {
    const workbook = XLSX.read(data, { type: 'array', cellDates: true });
    const worksheet = workbook.Sheets['Glava'];
    return {
        davcnaStevilka: parse_string(worksheet['B2']),
        obdobjeOd: parse_date(worksheet['B3']),
        obdobjeDo: parse_date(worksheet['B4']),
        vracilo: parse_boolean(worksheet['B5']),
        odbitniDelez: parse_boolean(worksheet['B6']),
        insolventniPostopek: parse_boolean(worksheet['B7']),
        odlocbaFurs: parse_boolean(worksheet['B8']),
        nacin: parse_method(worksheet['B9']),
        opomba: parse_string(worksheet['B10']),
    };
}


function parse_kir(data) {
    var parsed_data = [];
    const workbook = XLSX.read(data, { type: 'array', cellDates: true });
    const worksheet = workbook.Sheets['KIR'];

    var row = 10;
    while (worksheet['A' + row]) {
        parsed_data.push({
            zaporedna: parse_integer(worksheet['A' + row]),             // ZAPST
            datum_knjizenja: parse_date(worksheet['B' + row]),          // P2
            stevilka_racuna: parse_integer(worksheet['C' + row]),       // P3
            datum_racuna: parse_date(worksheet['D' + row]),             // P4
            obdobje: parse_string(worksheet['E' + row]),                // OBDOBJE
            nacin: parse_method(worksheet['F' + row]),                     // OBRAVNAVA
            podjetje: parse_string(worksheet['G' + row]),               // P5
            koda_drzave: parse_string(worksheet['H' + row]),            // P6
            davcna: parse_string(worksheet['I' + row]),                 // P6DS
            vrednost_brez_ddv: parse_float(worksheet['J' + row]),       // P7
            ddv_prejemnik: parse_float(worksheet['K' + row]),           // P8
            oproscene_dobave_slo: parse_float(worksheet['L' + row]),    // P9
            oproscene_dobave_eu: parse_float(worksheet['M' + row]),     // P10
            tristranske_dobave_eu: parse_float(worksheet['N' + row]),   // P11
            prodaja_na_daljavo: parse_float(worksheet['O' + row]),      // P12
            dobava_v_eu: parse_float(worksheet['P' + row]),             // P13
            ddv_22: parse_float(worksheet['Q' + row]),                  // P14
            ddv_9: parse_float(worksheet['R' + row]),                   // P15
            ddv_5: parse_float(worksheet['S' + row]),                   // P16
            prid_material_eu_22: parse_float(worksheet['T' + row]),     // P17
            prid_storitve_eu_22: parse_float(worksheet['U' + row]),     // P18
            prid_material_eu_9: parse_float(worksheet['V' + row]),      // P19
            prid_storitve_eu_9: parse_float(worksheet['W' + row]),      // P20
            prid_material_eu_5: parse_float(worksheet['X' + row]),      // P21
            prid_storitve_eu_5: parse_float(worksheet['Y' + row]),      // P22
            samoob_22: parse_float(worksheet['Z' + row]),               // P23
            samoob_9: parse_float(worksheet['AA' + row]),               // P24
            samoob_5: parse_float(worksheet['AB' + row]),               // P25
            samoob_uvoz: parse_float(worksheet['AC' + row]),            // P26
            dobave_zunaj_slo: parse_float(worksheet['AD' + row]),       // P27
            opombe: parse_string(worksheet['AE' + row]),                // P28
            samoprijava_obdobje: parse_string(worksheet['AF' + row]),   // OBDOBJE88
            samoprijava_davek: parse_string(worksheet['AG' + row]),     // DAVEK88
        });

        row++;
    }

    parsed_data.sort((a, b) => a.zaporedna - b.zaporedna);
    return parsed_data;
}

function parse_kpr(data) {
    var parsed_data = [];
    const workbook = XLSX.read(data, { type: 'array', cellDates: true });
    const worksheet = workbook.Sheets['KPR'];

    var row = 11;
    while (worksheet['A' + row]) {
        parsed_data.push({
            zaporedna: parse_integer(worksheet['A' + row]),             // ZAPST
            datum_knjizenja: parse_date(worksheet['B' + row]),          // P2
            stevilka_racuna: parse_integer(worksheet['C' + row]),       // P3
            datum_prejema: parse_date(worksheet['D' + row]),            // P4
            datum_racuna: parse_date(worksheet['E' + row]),             // P5
            obdobje: parse_string(worksheet['F' + row]),                // OBDOBJE
            nacin: parse_method(worksheet['G' + row]),                     // OBRAVNAVA
            podjetje: parse_string(worksheet['H' + row]),               // P6
            koda_drzave: parse_string(worksheet['I' + row]),            // P7
            davcna: parse_string(worksheet['J' + row]),                 // P7DS
            vrednost_brez_ddv: parse_float(worksheet['K' + row]),       // P8
            ddv_obrac_prejem: parse_float(worksheet['L' + row]),        // P9
            pridobitve_blaga_eu: parse_float(worksheet['M' + row]),     // P10
            pridobitve_storitev_eu: parse_float(worksheet['N' + row]),  // P11
            nepremicnine: parse_float(worksheet['O' + row]),            // P12
            osnovna_sredstva: parse_float(worksheet['P' + row]),        // P13
            oproscene_nabave: parse_float(worksheet['Q' + row]),        // P14
            oproscene_neprem: parse_float(worksheet['R' + row]),        // P15
            oproscena_oprema: parse_float(worksheet['S' + row]),        // P16
            ne_obije: parse_float(worksheet['T' + row]),                // P17
            ddv_22: parse_float(worksheet['U' + row]),                  // P18
            ddv_9: parse_float(worksheet['V' + row]),                   // P19
            ddv_5: parse_float(worksheet['W' + row]),                   // P20
            pavsal_8: parse_float(worksheet['X' + row]),                // P21
            opombe: parse_string(worksheet['Y' + row]),                 // P22
            samoprijava_obdobje: parse_string(worksheet['Z' + row]),    // OBDOBJE88
            samoprijava_davek: parse_string(worksheet['AA' + row]),     // DAVEK88
        });

        row++;
    }

    parsed_data.sort((a, b) => a.zaporedna - b.zaporedna);
    return parsed_data;
}

function validate_file(event, data) {
    // TODO validate tax id?
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
        display_reset();
        return false;
    }

    const workbook = XLSX.read(data, { type: 'array', cellDates: true });
    if (!workbook.SheetNames.includes("Glava") || !workbook.SheetNames.includes("KIR") || !workbook.SheetNames.includes("KPR")) {
        errorDiv.textContent = 'Pričakovani listi v Excel datoteki so: Glava, KIR, KPR.';
        display_reset();
        return false;
    }

    return true;
}

function generate_furs_json(header, kir, kpr) {
    // Note: This does not pass validation on eDavki and I can't figure out why. Keeping for
    // historical reasons. We are using XML instead.
    var furs_json = {
        Glava: {
            TaxPayerID: header.davcnaStevilka.value,
            // TODO do we need this?
            //TUJEC1: "AB",
            //TUJEC2: "string",
            // TODO How can we only get this once?
            OBDOBJE_OD: formatDate(header.obdobjeOd.value, "yyyy-MM-dd"),
            OBDOBJE_DO: formatDate(header.obdobjeDo.value, "yyyy-MM-dd"),
            KIR: kir.length > 0,
            KPR: kpr.length > 0,
            VRACILO: header.vracilo.value,
            ODBDELEZ: header.odbitniDelez.value,
            INSPOS: header.insolventniPostopek.value,
            PREDLODO: header.odlocbaFurs.value,
            OPOMBA: header.opomba.value,
        },
    };

    if (header.nacin.value) {
        furs_json.Glava.NACIN = header.nacin.value;
    }

    if (kir.length > 0) {
        furs_json.Lista_KIR = {};
        furs_json.Lista_KIR.KIR = [];
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
                OBDOBJE88: row.samoprijava_obdobje.value,
                DAVEK88: row.samoprijava_davek.value,
            });
        })
    }

    if (kpr.length > 0) {
        furs_json.Lista_KPR = {};
        furs_json.Lista_KPR.KPR = [];
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
                OBDOBJE88: row.samoprijava_obdobje.value,
                DAVEK88: row.samoprijava_davek.value,
            });
        });
    }

    return furs_json;
}

function generate_furs_xml(export_data) {
    const emptyValuesToOmit = [
        '.DDV_KIR_KPR.Lista_KIR.KIR.P6',
        '.DDV_KIR_KPR.Lista_KIR.KIR.P6DS',
        '.DDV_KIR_KPR.Lista_KIR.KIR.P28',
        '.DDV_KIR_KPR.Lista_KIR.KIR.OBDOBJE88',
        '.DDV_KIR_KPR.Lista_KIR.KIR.DAVEK88',

        '.DDV_KIR_KPR.Lista_KPR.KPR.P7',
        '.DDV_KIR_KPR.Lista_KPR.KPR.P22',
        '.DDV_KIR_KPR.Lista_KPR.KPR.OBDOBJE88',
        '.DDV_KIR_KPR.Lista_KPR.KPR.DAVEK88',
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

function generate_furs_files(data) {
    // Prepare JSON download
    const header = parse_header(data);
    const export_data = generate_furs_json(
        header,
        parse_kir(data),
        parse_kpr(data)
    );
    const jsonString = JSON.stringify(export_data, null, 2);
    const jsonBlob = new Blob([jsonString], { type: 'application/json' });
    const jsonFilename = `DDV_${header.davcnaStevilka.value.trim()}_${formatDate(header.obdobjeOd.value, "yyyyMM")}_${formatDate(header.obdobjeDo.value, "yyyyMM")}.json`;

    // Prepare XML download.
    const xmlString = generate_furs_xml(export_data);
    const xmlBlob = new Blob([xmlString], { type: 'application/xml' });
    const xmlFilename = `DDV_${header.davcnaStevilka.value.trim()}_${formatDate(header.obdobjeOd.value, "yyyyMM")}_${formatDate(header.obdobjeDo.value, "yyyyMM")}.xml`;

    downloadJson.href = URL.createObjectURL(jsonBlob);
    downloadJson.download = jsonFilename;

    downloadXml.href = URL.createObjectURL(xmlBlob);
    downloadXml.download = xmlFilename;

    // Create ZIP file
    const zip = new JSZip();
    zip.file(xmlFilename, xmlString);
    zip.generateAsync({ type: 'blob' }).then(function(zipBlob) {
        downloadZip.href = URL.createObjectURL(zipBlob);
        downloadZip.download = `DDV_${header.davcnaStevilka.value.trim()}_${formatDate(header.obdobjeOd.value, "yyyyMM")}_${formatDate(header.obdobjeDo.value, "yyyyMM")}.zip`;
        display_output();
    });

    // Create tables
    outputDiv.innerHTML = '';
    var heading = document.createElement('h3');
    heading.innerHTML = 'Glava';
    outputDiv.appendChild(heading);
    outputDiv.appendChild(generateTable(export_data.Glava, true));
    var heading = document.createElement('h3');
    heading.innerHTML = 'Knjiga izdanih računov';
    outputDiv.appendChild(heading);
    if ('Lista_KIR' in export_data) {
        outputDiv.appendChild(generateTable(export_data.Lista_KIR.KIR));
    }
    var heading = document.createElement('h3');
    heading.innerHTML = 'Knjiga prejetih računov';
    outputDiv.appendChild(heading);
    if ('Lista_KPR' in export_data) {
        outputDiv.appendChild(generateTable(export_data.Lista_KPR.KPR));
    }
}

excelFile.addEventListener('change', async (event) => {
    const data = await event.target.files[0].arrayBuffer();
    if (!validate_file(event, data)) {
        display_reset();
        return;
    }

    try {
        generate_furs_files(await event.target.files[0].arrayBuffer());
    } catch (error) {
        errorDiv.textContent = 'Error parsing file: ' + error.message;
        display_reset();
    }
});

resetForm.addEventListener('click', display_input);

function display_output() {
    fileUploadArea.style.display = 'none';
    downloadZip.style.display = 'inline-block';
    // downloadJson.style.display = 'inline-block';
    // downloadXml.style.display = 'inline-block';
    resetForm.style.display = 'inline-block';
}

function display_input() {
    downloadJson.style.display = 'none';
    downloadXml.style.display = 'none';
    downloadZip.style.display = 'none';
    resetForm.style.display = 'none';
    fileUploadArea.style.display = 'flex';
    outputDiv.innerHTML = '';
    errorDiv.innerHTML = '';
    excelFile.value = '';
}

function display_reset() {
    resetForm.style.display = 'inline-block';
    fileUploadArea.style.display = 'none';
    downloadJson.style.display = 'none';
    downloadXml.style.display = 'none';
    downloadZip.style.display = 'none';
}

// Drag and Drop functionality
function handleDragOver(event) {
    event.preventDefault();
    fileUploadArea.classList.add('dragover');
}

function handleDragLeave(event) {
    event.preventDefault();
    fileUploadArea.classList.remove('dragover');
}

function handleDrop(event) {
    event.preventDefault();
    fileUploadArea.classList.remove('dragover');

    const files = event.dataTransfer.files;
    if (files.length > 0) {
        const file = files[0];
        const validTypes = [
            'application/vnd.ms-excel',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        ];

        if (!validTypes.includes(file.type)) {
            showError('Omogočene so samo Excel datoteke (.xls, .xlsx)');
            return;
        }

        excelFile.files = files;
        updateFileName(file.name);
        excelFile.dispatchEvent(new Event('change'));
    }
}

// File name display
function updateFileName(name) {
    fileName.textContent = `Izbrana datoteka: ${name}`;
    fileName.style.display = 'block';
}

// Error handling
function showError(message) {
    errorText.textContent = message;
    errorDiv.style.display = 'flex';
    errorDiv.classList.add('fade-in');
}

function hideError() {
    errorDiv.style.display = 'none';
    errorText.textContent = '';
}

// Loading states
function showLoading(element, text = 'Obdelovanje...') {
    element.innerHTML = `<span class="loading"></span> ${text}`;
    element.disabled = true;
}

function hideLoading(element, originalText) {
    element.innerHTML = originalText;
    element.disabled = false;
}

// Update table generation to include section titles
function generate_furs_files(data) {
    showLoading(downloadZip, 'Ustvarjam datoteke...');

    // Prepare JSON download
    const header = parse_header(data);
    const export_data = generate_furs_json(
        header,
        parse_kir(data),
        parse_kpr(data)
    );
    const jsonString = JSON.stringify(export_data, null, 2);
    const jsonBlob = new Blob([jsonString], { type: 'application/json' });
    const jsonFilename = `DDV_${header.davcnaStevilka.value.trim()}_${formatDate(header.obdobjeOd.value, "yyyyMM")}_${formatDate(header.obdobjeDo.value, "yyyyMM")}.json`;

    // Prepare XML download.
    const xmlString = generate_furs_xml(export_data);
    const xmlBlob = new Blob([xmlString], { type: 'application/xml' });
    const xmlFilename = `DDV_${header.davcnaStevilka.value.trim()}_${formatDate(header.obdobjeOd.value, "yyyyMM")}_${formatDate(header.obdobjeDo.value, "yyyyMM")}.xml`;

    downloadJson.href = URL.createObjectURL(jsonBlob);
    downloadJson.download = jsonFilename;

    downloadXml.href = URL.createObjectURL(xmlBlob);
    downloadXml.download = xmlFilename;

    // Create ZIP file
    const zip = new JSZip();
    zip.file(xmlFilename, xmlString);
    zip.generateAsync({ type: 'blob' }).then(function(zipBlob) {
        downloadZip.href = URL.createObjectURL(zipBlob);
        downloadZip.download = `DDV_${header.davcnaStevilka.value.trim()}_${formatDate(header.obdobjeOd.value, "yyyyMM")}_${formatDate(header.obdobjeDo.value, "yyyyMM")}.zip`;
        hideLoading(downloadZip, '📦 Prenesi ZIP');
        display_output();
    });

    // Create tables with improved structure
    outputDiv.innerHTML = '';

    // Company info section
    const companySection = document.createElement('section');
    companySection.className = 'output-section card';
    const companyTitle = document.createElement('h3');
    companyTitle.className = 'section-title';
    companyTitle.textContent = '❖ Osnovni podatki';
    companySection.appendChild(companyTitle);
    companySection.appendChild(generateTable(export_data.Glava, true));
    outputDiv.appendChild(companySection);

    // KIR section
    if ('Lista_KIR' in export_data && export_data.Lista_KIR.KIR.length > 0) {
        const kirSection = document.createElement('section');
        kirSection.className = 'output-section card';
        const kirTitle = document.createElement('h3');
        kirTitle.className = 'section-title';
        kirTitle.textContent = '💰 Knjiga izdanih računov (KIR)';
        kirSection.appendChild(kirTitle);
        kirSection.appendChild(generateTable(export_data.Lista_KIR.KIR));
        outputDiv.appendChild(kirSection);
    }

    // KPR section
    if ('Lista_KPR' in export_data && export_data.Lista_KPR.KPR.length > 0) {
        const kprSection = document.createElement('section');
        kprSection.className = 'output-section card';
        const kprTitle = document.createElement('h3');
        kprTitle.className = 'section-title';
        kprTitle.textContent = '📥 Knjiga prejetih računov (KPR)';
        kprSection.appendChild(kprTitle);
        kprSection.appendChild(generateTable(export_data.Lista_KPR.KPR));
        outputDiv.appendChild(kprSection);
    }
}

// Event listeners
excelFile.addEventListener('change', async (event) => {
    const file = event.target.files[0];
    if (!file) return;

    updateFileName(file.name);
    hideError();

    const validTypes = [
        'application/vnd.ms-excel',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    ];

    if (!validTypes.includes(file.type)) {
        showError('Omogočene so samo Excel datoteke (.xls, .xlsx)');
        return;
    }

    try {
        const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true });
        if (!workbook.SheetNames.includes("Glava") || !workbook.SheetNames.includes("KIR") || !workbook.SheetNames.includes("KPR")) {
            showError('Pričakovani listi v Excel datoteki so: Glava, KIR, KPR');
            display_reset();
            return;
        }
        generate_furs_files(await file.arrayBuffer());
    } catch (error) {
        showError('Napaka pri obdelavi datoteke: ' + error.message);
        display_reset();
    }
});

resetForm.addEventListener('click', () => {
    display_input();
    fileName.style.display = 'none';
    hideError();
    excelFile.value = '';
});

fileUploadArea.addEventListener('dragover', handleDragOver);
fileUploadArea.addEventListener('dragleave', handleDragLeave);
fileUploadArea.addEventListener('drop', handleDrop);
