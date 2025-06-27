import { isDefaultClause, transpileModule, unescapeLeadingUnderscores } from 'typescript';
import * as XLSX from 'xlsx';

const kprFile = document.getElementById('kprFile');
const kirFile = document.getElementById('kirFile');
const outputDiv = document.getElementById('output');
const errorDiv = document.getElementById('error');
const downloadJson = document.getElementById('downloadJson');

const davcnaStevilka = document.getElementById('davcnaStevilka');
const zahtevamVracilo = document.getElementById('zahtevamVracilo');
const izracunavamOdbitniDelez = document.getElementById('izracunavamOdbitniDelez');
const nacin = document.getElementById('nacin');
const opomba = document.getElementById('opomba');
const obdobjeOd = document.getElementById('obdobjeOd');
const obdobjeDo = document.getElementById('obdobjeDo');

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
            opombe: parse_float(worksheet['AC' + row]),                 // P28
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
            opombe: parse_float(worksheet['W' + row]),                  // P22
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
    // TODO validate tax id?
    if (kir == undefined || kpr == undefined || davcnaStevilka.value.trim() == '' || obdobjeOd.value.trim() == '' || obdobjeDo.value.trim() == '') {
        downloadJson.style.display = 'none';
        furs_json = undefined;
        return;
    }

    furs_json = {
        Glava: {
            TaxPayerID: davcnaStevilka.value.trim(),
            // TODO do we need this?
            //TUJEC1: "AB",
            //TUJEC2: "string",
            // TODO How can we only get this once?
            OBDOBJE_OD: (new Date(obdobjeOd.value)).toISOString(),
            OBDOBJE_DO: (new Date(obdobjeDo.value)).toISOString(),
            // TODO se zgodi da je kateri prazen?
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
            P2: row.datum_knjizenja.value.toISOString(),
            P3: row.stevilka_racuna.value,
            P4: row.datum_racuna.value.toISOString(),
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
            P29: row.opombe.value,
            OBRAVNAVA: row.nacin.value,
        });
    })

    kpr.forEach((row) => {
        furs_json.Lista_KPR.KPR.push({
            ZAPST: row.zaporedna.value,
            OBDOBJE: row.obdobje.value,
            P2: row.datum_knjizenja.value.toISOString(),
            P3: row.stevilka_racuna.value,
            P4: row.datum_prejema.value.toISOString(),
            P5: row.datum_racuna.value.toISOString(),
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

    // Prepare JSON download
    const jsonString = JSON.stringify(furs_json, null, 2);
    const blob = new Blob([jsonString], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    downloadJson.href = url;
    downloadJson.style.display = 'inline-block';
}

kprFile.addEventListener('change', async (event) => {
    if (!validate_file(event)) {
        return;
    }

    try {
        kpr = parse_kpr(await event.target.files[0].arrayBuffer());
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
        kir = parse_kir(await event.target.files[0].arrayBuffer());
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

davcnaStevilka.addEventListener('change', async (event) => {
    generate_furs_json();
});
zahtevamVracilo.addEventListener('change', async (event) => {
    generate_furs_json();
});
izracunavamOdbitniDelez.addEventListener('change', async (event) => {
    generate_furs_json();
});
nacin.addEventListener('change', async (event) => {
    generate_furs_json();
});
opomba.addEventListener('change', async (event) => {
    generate_furs_json();
});
obdobjeOd.addEventListener('change', async (event) => {
    generate_furs_json();
});
obdobjeDo.addEventListener('change', async (event) => {
    generate_furs_json();
});
