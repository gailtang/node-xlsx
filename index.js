const XLSX = require('xlsx');
const m = require('moment');
let src = XLSX.readFile('source.xlsx');
let dst = XLSX.readFile('destination.xlsx');

function directMutate(mapping) {
    let src_sheet = src.Sheets[mapping.src_sheet];
    if (!src_sheet) {
        console.log('failed to get source sheet from mapping:', mapping);
        process.exit(1)
    }

    let dst_sheet = dst.Sheets[mapping.dst_sheet];
    if (!dst_sheet) {
        console.log('failed to get destination sheet from mapping:', mapping);
        process.exit(1)
    }
    let src_name2Col = getColNames(src.Sheets[mapping.src_sheet]);
    let src_name2Row = getRowNames(src.Sheets[mapping.src_sheet]);
    let dst_name2Col = getColNames(dst.Sheets[mapping.dst_sheet]);
    let dst_name2Row = getRowNames(dst.Sheets[mapping.dst_sheet]);

    for (let i = 0; i < mapping.sections.length; i++) {
        let section = mapping.sections[i];
        for (let j = 0; j < mapping.mappings.cols.length; j++) {
            let srcColName = mapping.mappings.cols[j].src;
            let dstColName = mapping.mappings.cols[j].dst;
            if (dstColName == 'default')
                dstColName = mapping.src_sheet;
            for (let k = 0; k < mapping.mappings.rows.length; k++) {
                let srcRowName = mapping.mappings.rows[k].src;
                let dstRowName = mapping.mappings.rows[k].dst;
                try {
                    let r_src = src_name2Row[srcRowName][section];
                    let c_src = src_name2Col[srcColName];
                    let r_dst = dst_name2Row[dstRowName][section];
                    let c_dst = dst_name2Col[dstColName];
                    let src_cell_name = XLSX.utils.encode_cell({ c: c_src, r: r_src });
                    let dst_cell_name = XLSX.utils.encode_cell({ c: c_dst, r: r_dst });
                    dst_sheet[dst_cell_name] = src_sheet[src_cell_name];
                    console.log(
                        `  ${src_cell_name}(${srcColName},${srcRowName},${section})=>${dst_cell_name}(${dstColName},${dstRowName},${section})  value:${src_sheet[src_cell_name].w}`
                        //                    `Copy sheet:${mapping.src_sheet} cell: ${src_cell_name}/[${srcColName},${srcRowName}, ${srcSection}]` +
                        //                    ` => sheet:${mapping.dst_sheet} cell: ${dst_cell_name}/[${dstColName},${dstRowName}, ${dstSection}]` +
                        //                    ` value:${src_sheet[src_cell_name].v}`
                    );

                } catch (e) {
                    console.log(e);
                    console.log('failed to get cells:');
                    process.exit(1);
                }
            }
        }
    };

}

function rtocMutate(mapping) {
    let src_sheet = src.Sheets[mapping.src_sheet];
    if (!src_sheet) {
        console.log('failed to get source sheet from mapping:', mapping);
        process.exit(1)
    }

    let dst_sheet = dst.Sheets[mapping.dst_sheet];
    if (!dst_sheet) {
        console.log('failed to get destination sheet from mapping:', mapping);
        process.exit(1)
    }

    let src_name2Col = getColNames(src.Sheets[mapping.src_sheet]);
    let src_name2Row = getRowNames(src.Sheets[mapping.src_sheet]);
    let dst_name2Col = getColNames(dst.Sheets[mapping.dst_sheet]);
    let dst_name2Row = getRowNames(dst.Sheets[mapping.dst_sheet]);

    let dstColName = mapping.default_dst_col;
    if (dstColName == 'default')
        dstColName = mapping.src_sheet;
    let c_dst = dst_name2Col[dstColName];
    let srcSection = mapping.default_src_section;

    for (let i = 0; i < mapping.mappings.length; i++) {
        let m = mapping.mappings[i];
        let srcRowName = m.src.row;
        let r_src = src_name2Row[srcRowName][srcSection];
        for (let j = 0; j < m.src.cols.length; j++) {
            try {
                let srcColName = m.src.cols[j];
                let dstRowName = m.dst.cols[j];
                let dstSection = m.dst.section;
                let c_src = src_name2Col[srcColName];
                let r_dst = dst_name2Row[dstRowName][dstSection];
                let src_cell_name = XLSX.utils.encode_cell({ c: c_src, r: r_src });
                let dst_cell_name = XLSX.utils.encode_cell({ c: c_dst, r: r_dst });
                dst_sheet[dst_cell_name] = src_sheet[src_cell_name];
                console.log(
                    `  ${src_cell_name}(${srcColName},${srcRowName},${srcSection})=>${dst_cell_name}(${dstColName},${dstRowName},${dstSection})  value:${src_sheet[src_cell_name].w}`
                    //                    `Copy sheet:${mapping.src_sheet} cell: ${src_cell_name}/[${srcColName},${srcRowName}, ${srcSection}]` +
                    //                    ` => sheet:${mapping.dst_sheet} cell: ${dst_cell_name}/[${dstColName},${dstRowName}, ${dstSection}]` +
                    //                    ` value:${src_sheet[src_cell_name].v}`
                );

            } catch (e) {
                console.log(e);
                console.log('failed to get cells:');
                process.exit(1);
            }
        }
    }
}

function getColNames(sheet) {
    let ref = sheet['!ref'];
    let name2Col = {};
    let maxCell = ref.split(':')[1];
    let decodedCell = XLSX.utils.decode_cell(maxCell);
    let maxCol = decodedCell.c;
    let maxRow = decodedCell.r;

    function getHeaderRow() {
        for (let i = 0; i <= maxRow; i++) {
            let count = 0;
            for (let j = 0; j <= maxCol; j++) {
                let cell = sheet[XLSX.utils.encode_cell({ c: j, r: i })];
                if (cell)
                    count++;
            }
            if (count / maxCol > 0.5)
                return i;
        }
        return -1;
    }
    let headerRow = getHeaderRow();
    if (headerRow >= 0) {
        for (let i = 0; i <= maxCol; i++) {
            let cell = sheet[XLSX.utils.encode_cell({ c: i, r: headerRow })];
            if (cell && typeof cell.v == 'string')
                name2Col[cell.v.trim()] = i;
        }
    }
    // console.log(name2Col);
    return name2Col;
}


function getRowNames(sheet) {
    let name2Row = {};
    let ref = sheet['!ref'];
    let maxCell = ref.split(':')[1];
    let decodedCell = XLSX.utils.decode_cell(maxCell);
    let maxRow = decodedCell.r;
    let maxCol = decodedCell.c;

    function getHeaderCol() {
        for (let i = 0; i <= maxCol; i++) {
            let count = 0;
            for (let j = 0; j <= maxRow; j++) {
                let cell = sheet[XLSX.utils.encode_cell({ c: i, r: j })];
                if (cell)
                    count++;
            }
            if (count / maxCol > 0.5)
                return i;
        }
        return -1;
    }

    let headerCol = getHeaderCol();
    for (let i = 0; i <= maxRow; i++) {
        let cell = sheet[XLSX.utils.encode_cell({ c: headerCol, r: i })];
        if (cell) {
            let key = cell.v.trim();
            if (typeof name2Row[key] == 'undefined')
                name2Row[key] = [];
            name2Row[key].push(i)
        }
    }
    // console.log(name2Row);
    return name2Row;
}

if (process.argv.length == 3) {

} else {
    const mappings = require('./mappings.json');
    mappings.forEach((element, i) => {
        console.log(`Process rule ${i}, type:${element.type}, source sheet: ${element.src_sheet}, destination sheet: ${element.dst_sheet}`);
        switch (element.type) {
            case 'DIRECT':
                directMutate(element);
                break;
            case 'RTOC':
                rtocMutate(element);
                break;
            default:
                break;
        }
    });
}

/* DO SOMETHING WITH workbook HERE */
let t = m().format('YYYY-MM-DD_HHmmss');
let output = 'output_' + t + '.xlsx';
XLSX.writeFile(dst, output);
console.log('Done, output is written to file:', output);

console.log('Press any key to exit');

/* process.stdin.setRawMode(true);
process.stdin.resume();
process.stdin.on('data', process.exit.bind(process, 0)); */