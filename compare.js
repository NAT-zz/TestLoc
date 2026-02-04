const ExcelJS = require('exceljs')
const fs = require('fs')
const path = require('path')

const FILE_CHINH = 'main.xlsx'
const DIR_PHU = './CSDLYDUOC'

const COL_CHINH = 'Số giấy phép hoạt động'
const COL_PHU = 'Số GPHĐ'

function normalize(val) {
    if (!val) return null
    if (typeof val === 'string') return val.trim()
    if (val.text) return val.text.trim()
    if (val.richText) return val.richText.map(t => t.text).join('').trim()
    return String(val).trim()
}

/**
 * Map<GPHĐ, Array<{ file, extra }>>
 * extra = giá trị cột 3
 */
async function readSubFile(filePath, gphdMap) {
    const wb = new ExcelJS.Workbook()
    await wb.xlsx.readFile(filePath)
    const ws = wb.worksheets[0]

    const fileName = path.basename(filePath)

    let colIndex
    ws.getRow(4).eachCell((cell, col) => {
        if (normalize(cell.value) === COL_PHU) colIndex = col
    })

    if (!colIndex) {
        console.warn(`⚠️ Không có cột Số GPHĐ trong ${fileName}`)
        return
    }

    ws.eachRow((row, idx) => {
        if (idx <= 4) return

        const val = normalize(row.getCell(colIndex).value)
        if (!val) return

        const extra = normalize(row.getCell(3).value) // ⭐ CỘT 3

        if (!gphdMap.has(val)) {
            gphdMap.set(val, [])
        }

        gphdMap.get(val).push({
            file: fileName,
            extra: extra
        })
    })
}

async function run() {
    /** ===== 1. ĐỌC FILE PHỤ ===== */
    const gphdMap = new Map()

    const files = fs.readdirSync(DIR_PHU)
        .filter(f => f.endsWith('.xlsx'))

    for (const file of files) {
        const fullPath = path.join(DIR_PHU, file)
        console.log('📄 Đang đọc:', fullPath)
        await readSubFile(fullPath, gphdMap)
    }

    console.log('🔎 Tổng GPHĐ (unique):', gphdMap.size)

    /** ===== 2. FILE CHÍNH ===== */
    const wb = new ExcelJS.Workbook()
    await wb.xlsx.readFile(FILE_CHINH)
    const ws = wb.worksheets[1]

    let colChinh
    ws.getRow(1).eachCell((c, i) => {
        if (normalize(c.value) === COL_CHINH) colChinh = i
    })
    if (!colChinh) throw new Error('Không tìm thấy cột Số giấy phép hoạt động')

    ws.eachRow((row, idx) => {
        if (idx === 1) return

        const cell = row.getCell(colChinh)
        const val = normalize(cell.value)
        if (!val) return
        if (!gphdMap.has(val)) return

        const target = cell.isMerged ? cell.master : cell
        const newStyle = JSON.parse(JSON.stringify(target.style || {}))
        newStyle.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF6AFF00' }
        }

        target.style = newStyle

        // 🔥 MATCH → XOÁ
        gphdMap.delete(val)
    })

    /** ===== 3. GHI CÁC GPHĐ CÒN LẠI RA EXCEL ===== */
    if (gphdMap.size > 0) {
        const logWb = new ExcelJS.Workbook()
        const logWs = logWb.addWorksheet('GPHĐ_KHÔNG_TRÙNG')

        // Header
        logWs.addRow(['Tên file', 'Số GPHĐ', 'Cơ sở'])

        // Style header
        logWs.getRow(1).font = { bold: true }
        logWs.columns = [
            { width: 25 },
            { width: 25 },
            { width: 40 }
        ]

        // Data
        for (const [val, rows] of gphdMap.entries()) {
            for (const r of rows) {
                logWs.addRow([
                    r.file,
                    val,
                    r.extra ?? '(trống)'
                ])
            }
        }

        await logWb.xlsx.writeFile('GPHD_KHONG_TRUNG.xlsx')
        console.log('📄 Đã ghi file GPHD_KHONG_TRUNG.xlsx')
    } else {
        console.log('✅ Tất cả GPHĐ đều đã được đối soát')
    }

    await wb.xlsx.writeFile('output_SAFE.xlsx')
    console.log('\n✅ Hoàn tất')
}

run().catch(console.error)
