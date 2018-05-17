// XSLX importer for OTVS
// Writes an SQL INSERT statement to stdout
//
// Usage:
// node index.js <file>.xlsx <otvs_meta_id>
//
// Notes:
// - Turn date cells into text before
//

const moment = require("moment");
const xlsx = require("xlsx");

const map = {
  id: (row, i) => i + 1,
  course: "Teilnehmer::RKurs",
  coursetitle: "Teilnehmer::RKurstitel",
  last: "Nachname",
  first: "Vorname",
  email: "#EMail",
  birthdate: "Geburtsdatum",
  nationality: "#Staatsangehörigkeit",
  sex: "Geschlecht",
  street: "#Strasse_Nr_PF",
  zip: "#PLZ",
  city: "#Ort",
  country: "#Staat",
  phone: "#Telefon",
  mobile: "#Mobil",
  username: () => null,
  state: "#Bundesland",
  emname: () => null,
  emaddress: () => null,
  emphone: () => null,
  emmobile: () => null,
  locked: () => false,
  is_verified: () => false,
  email_verification_code: () => null,
  diet: () => null,
  allergies: () => null,
  meds: () => null,
  insurancetype: () => null,
  insurancecompany: () => null,
  schoolname: {
    fields: [
      "DF_Schuladresse::Adresszeile_1",
      "DF_Schuladresse::Adresszeile_2"
    ],
    joiner: " "
  },
  schooladdress: "DF_Schuladresse::Straße_Nr_PF",
  schoolzip: "DF_Schuladresse::PLZ",
  schoolcity: "DF_Schuladresse::Ort",
  schoolstate: "DF_Schuladresse::Bundesland",
  schoolcountry: "DF_Schuladresse::Staat",
  schoolgrade: "Teilnehmer::Klasse",
  schoolprofile: [
    "Teilnehmer::LS1a_LK_Fach_DEU",
    "Teilnehmer::LS1b_LK_Fach_DEU",
    "Teilnehmer::LS1c_LK_Fach_DEU",
    "Teilnehmer::LS1d_LK_Fach_DEU",
    "Teilnehmer::LS1e_LK_Fach_DEU"
  ],
  musicinstrument: [
    "Teilnehmer::LS5a_Instrument",
    "Teilnehmer::LS5b_Instrument",
    "Teilnehmer::LS5c_Instrument",
    "Teilnehmer::LS5d_Instrument",
    "Teilnehmer::LS5e_Instrument"
  ],
  musicvoice: ["Teilnehmer::LS6a_Stimmlage", "Teilnehmer::LS6b_Stimmlage"],
  mothertongue: [
    "Teilnehmer::LS3a_Muttersprache",
    "Teilnehmer::LS3b_Muttersprache"
  ],
  foreignlang: [
    "Teilnehmer::LS4a_Fremdsprache",
    "Teilnehmer::LS4b_Fremdsprache",
    "Teilnehmer::LS4c_Fremdsprache",
    "Teilnehmer::LS4d_Fremdsprache"
  ],
  arrival: () => null
};

const id = a => a;

// Source: http://stackoverflow.com/a/7760578
const mysql_real_escape_string = str =>
  str.replace(/[\0\x08\x09\x1a\n\r"'\\\%]/g, function(char) {
    switch (char) {
      case "\0":
        return "\\0";
      case "\x08":
        return "\\b";
      case "\x09":
        return "\\t";
      case "\x1a":
        return "\\z";
      case "\n":
        return "\\n";
      case "\r":
        return "\\r";

      case '"':
        return '\\"';
      case "'":
        return "\\'";
      case "\\":
        return "\\\\";
      case "%":
        return "\\%";
    }
  });

const parseSheet = function(sheet) {
  function encode_col(col) {
    let s = "";
    for (++col; col; col = Math.floor((col - 1) / 26))
      s = String.fromCharCode((col - 1) % 26 + 65) + s;
    return s;
  }
  function encode_row(row) {
    return `${row + 1}`;
  }
  function encode_cell(cell) {
    return encode_col(cell.c) + encode_row(cell.r);
  }
  function decode_col(c) {
    let d = 0,
      i = 0;
    for (; i !== c.length; ++i) d = 26 * d + c.charCodeAt(i) - 64;
    return d - 1;
  }
  function decode_row(rowstr) {
    return Number(rowstr) - 1;
  }
  function split_cell(cstr) {
    return cstr.replace(/(\$?[A-Z]*)(\$?[0-9]*)/, "$1,$2").split(",");
  }
  function decode_cell(cstr) {
    const splt = split_cell(cstr);
    return { c: decode_col(splt[0]), r: decode_row(splt[1]) };
  }
  function decode_range(range) {
    const x = range.split(":").map(decode_cell);
    return { s: x[0], e: x[x.length - 1] };
  }

  const range = decode_range(sheet["!ref"]);
  const output = [];

  for (
    let i = range.s.r, end = range.e.r, asc = range.s.r <= end;
    asc ? i <= end : i >= end;
    asc ? i++ : i--
  ) {
    output[i] = [];
    for (
      let j = range.s.c, end1 = range.e.c, asc1 = range.s.c <= end1;
      asc1 ? j <= end1 : j >= end1;
      asc1 ? j++ : j--
    ) {
      output[i][j] = __guard__(sheet[encode_cell({ c: j, r: i })], x => x.v);
    }
  }

  return output;
};

{
  if (process.argv.length < 4) {
    console.error("Usage: node index.js <file>.xlsx <otvs_meta_id>");
    process.exit(1);
  }

  const file = xlsx.readFile(process.argv[2]);
  const sheet = file.Sheets[file.SheetNames[0]];
  const rows = parseSheet(sheet);

  const headers = {};

  for (let i = 0; i < rows[0].length; i++) {
    const field = rows[0][i];
    headers[field] = i;
  }

  const iterable = rows.slice(1);
  for (let i = 0; i < iterable.length; i++) {
    const row = iterable[i];
    const entry = {};

    for (const outputField in map) {
      var inputField = map[outputField];
      entry[outputField] = (() => {
        if (
          Object.prototype.toString.apply(inputField) === "[object Function]"
        ) {
          let left;
          return (left = inputField(row, i, headers)) != null ? left : "";
        } else if (
          Object.prototype.toString.apply(inputField) === "[object Array]"
        ) {
          for (const field of Array.from(inputField)) {
            if (headers[field] == null) {
              throw `no field '${field}' found`;
            }
          }
          return inputField
            .map(a => row[headers[a]])
            .filter(id)
            .join(", ");
        } else if (
          Object.prototype.toString.apply(inputField) === "[object Object]"
        ) {
          for (const field of Array.from(inputField.fields)) {
            if (headers[field] == null) {
              throw `no field '${field}' found`;
            }
          }
          return inputField.fields
            .map(a => row[headers[a]])
            .filter(id)
            .join(inputField.joiner);
        } else {
          if (headers[inputField] == null) {
            throw `no field '${inputField}' found`;
          }
          return row[headers[inputField]];
        }
      })();
    }

    console.log(
      `INSERT INTO wp_otvs(\`key\`, \`value\`, created_at, updated_at) VALUES("entry-${
        process.argv[3]
      }-${entry.id}", "${mysql_real_escape_string(
        JSON.stringify(entry)
      )}", ${(Date.now() / 1000) | 0}, ${(Date.now() / 1000) | 0});`
    );
  }
}

function __guard__(value, transform) {
  return typeof value !== "undefined" && value !== null
    ? transform(value)
    : undefined;
}
