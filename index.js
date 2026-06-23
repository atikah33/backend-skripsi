import express from 'express'
import cors from 'cors'
import multer from 'multer'
import dotenv from 'dotenv'
import fs from 'fs'
import xlsx from 'xlsx'
import mammoth from 'mammoth'
import { PDFParse } from 'pdf-parse'
import { createClient } from '@supabase/supabase-js'

dotenv.config()

const app = express()
app.use(cors())
app.use(express.json())

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_KEY,
)

const upload = multer({ dest: 'uploads/' })

const parseExcel = (filePath) => {
  const workbook = xlsx.readFile(filePath)
  const sheet = workbook.Sheets[workbook.SheetNames[0]]
  return xlsx.utils.sheet_to_json(sheet)
}

const parseDocx = async (filePath) => {
  const result = await mammoth.extractRawText({ path: filePath })
  return result.value
}

const parsePdf = async (filePath) => {
  const parser = new PDFParse({ url: filePath })
  const result = await parser.getText()
  return result.text
}

const cleanPdfLine = (line) => {
  const text = String(line || '').trim()

  if (!text) return ''

  // buang footer/header umum PDF
  if (/^-+\s*\d+\s+of\s+\d+\s*-+$/i.test(text)) return ''
  if (/^page\s+\d+\s+of\s+\d+$/i.test(text)) return ''
  if (/^\d+\s*\/\s*\d+$/.test(text)) return ''
  if (/^-+\s*page\s+\d+\s*-+$/i.test(text)) return ''

  return text
}

const parseQuestionText = (text) => {
  const normalized = String(text || '')
    .replace(/\r/g, '\n')
    .replace(/([^\n])(\d+\.\s*)/g, '$1\n\n$2')
    .replace(/([^\n])(A[\.\)]\s*)/g, '$1\n$2')
    .replace(/([^\n])(B[\.\)]\s*)/g, '$1\n$2')
    .replace(/([^\n])(C[\.\)]\s*)/g, '$1\n$2')
    .replace(/([^\n])(D[\.\)]\s*)/g, '$1\n$2')
    .replace(/([^\n])((Jawaban|Kunci)\s*:)/gi, '$1\n$2')
    .replace(/([^\n])(Bobot\s*:)/gi, '$1\n$2')
    .replace(/([^\n])(Tipe\s*:)/gi, '$1\n$2')
    .replace(/\n{3,}/g, '\n\n')
    .trim()

  const blocks = normalized
    .split(/\n\s*\n|(?=^\d+\.\s*)/m)
    .map((item) => item.trim())
    .filter(Boolean)

  const questions = []

  blocks.forEach((block) => {
    const lines = block
      .split('\n')
      .map(cleanPdfLine)
      .filter(Boolean)

    if (!lines.length) return

    const firstLine = lines[0]

    // skip block yang jelas bukan soal
    if (/^-+\s*\d+\s+of\s+\d+\s*-+$/i.test(firstLine)) return
    if (/^page\s+\d+\s+of\s+\d+$/i.test(firstLine)) return
    if (/^\d+\s*\/\s*\d+$/.test(firstLine)) return

    const looksLikeQuestion =
      /^\d+\.\s*/.test(firstLine) ||
      firstLine.includes('?') ||
      lines.some((line) => /^A[\.\)]\s*/i.test(line))

    if (!looksLikeQuestion) return

    let pertanyaan = firstLine.replace(/^\d+\.\s*/, '').trim()
    let tipe = ''
    let opsi_a = ''
    let opsi_b = ''
    let opsi_c = ''
    let opsi_d = ''
    let jawaban = ''
    let bobot = 10

    const tambahanPertanyaan = []

    lines.slice(1).forEach((line) => {
      if (/^A[\.\)]\s*/i.test(line)) {
        opsi_a = line.replace(/^A[\.\)]\s*/i, '').trim()
      } else if (/^B[\.\)]\s*/i.test(line)) {
        opsi_b = line.replace(/^B[\.\)]\s*/i, '').trim()
      } else if (/^C[\.\)]\s*/i.test(line)) {
        opsi_c = line.replace(/^C[\.\)]\s*/i, '').trim()
      } else if (/^D[\.\)]\s*/i.test(line)) {
        opsi_d = line.replace(/^D[\.\)]\s*/i, '').trim()
      } else if (/^(jawaban|kunci)\s*:/i.test(line)) {
        jawaban = line.replace(/^(jawaban|kunci)\s*:\s*/i, '').trim()
      } else if (/^bobot\s*:/i.test(line)) {
        bobot = Number(line.replace(/^bobot\s*:\s*/i, '').trim() || 10)
      } else if (/^tipe\s*:/i.test(line)) {
        tipe = line.replace(/^tipe\s*:\s*/i, '').toLowerCase().trim()
      } else {
        tambahanPertanyaan.push(line)
      }
    })

    if (tambahanPertanyaan.length && !opsi_a && !opsi_b && !opsi_c && !opsi_d) {
      pertanyaan = `${pertanyaan}\n${tambahanPertanyaan.join('\n')}`
    }

    let finalTipe = 'essay'

    if (tipe === 'pg' || tipe === 'pilihan ganda') {
      finalTipe = 'pg'
    } else if (tipe === 'essay' || tipe === 'esai') {
      finalTipe = 'essay'
    } else if (opsi_a || opsi_b || opsi_c || opsi_d) {
      finalTipe = 'pg'
    }

    questions.push({
      pertanyaan,
      tipe_soal: finalTipe,
      opsi_a: finalTipe === 'pg' ? opsi_a : null,
      opsi_b: finalTipe === 'pg' ? opsi_b : null,
      opsi_c: finalTipe === 'pg' ? opsi_c : null,
      opsi_d: finalTipe === 'pg' ? opsi_d : null,
      jawaban_benar: jawaban || null,
      bobot: Number(bobot || 10),
      urutan: questions.length + 1,
    })
  })

  return questions
}

const getExcelValue = (row, keys, defaultValue = '') => {
  for (const key of keys) {
    if (row[key] !== undefined && row[key] !== null && row[key] !== '') {
      return row[key]
    }
  }

  return defaultValue
}

const parseExcelRowsToQuestions = (rows) => {
  return rows.map((row, index) => {
    const pertanyaan = getExcelValue(row, [
      'pertanyaan',
      'Pertanyaan',
      'PERTANYAAN',
      'soal',
      'Soal',
      'SOAL',
    ])

    if (!pertanyaan) {
      throw new Error(`Baris ${index + 2}: pertanyaan/soal wajib diisi.`)
    }

    const tipeRaw = String(
      getExcelValue(row, [
        'tipe_soal',
        'tipe soal',
        'Tipe_Soal',
        'Tipe Soal',
        'TIPE_SOAL',
        'TIPE SOAL',
        'tipe',
        'Tipe',
      ]),
    )
      .toLowerCase()
      .trim()

    const opsiA = getExcelValue(row, ['opsi_a', 'opsi a', 'Opsi_A', 'Opsi A', 'A'])
    const opsiB = getExcelValue(row, ['opsi_b', 'opsi b', 'Opsi_B', 'Opsi B', 'B'])
    const opsiC = getExcelValue(row, ['opsi_c', 'opsi c', 'Opsi_C', 'Opsi C', 'C'])
    const opsiD = getExcelValue(row, ['opsi_d', 'opsi d', 'Opsi_D', 'Opsi D', 'D'])

    const jawaban = getExcelValue(row, [
      'jawaban_benar',
      'jawaban benar',
      'Jawaban_Benar',
      'Jawaban Benar',
      'JAWABAN_BENAR',
      'JAWABAN BENAR',
      'kunci',
      'Kunci',
      'kunci_jawaban',
      'Kunci Jawaban',
    ])

    const bobot = getExcelValue(row, ['bobot', 'Bobot', 'BOBOT'], 10)

    let finalTipe = 'pg'

    if (tipeRaw === 'essay' || tipeRaw === 'esai') {
      finalTipe = 'essay'
    } else if (tipeRaw === 'pg' || tipeRaw === 'pilihan ganda') {
      finalTipe = 'pg'
    } else {
      finalTipe = opsiA || opsiB || opsiC || opsiD ? 'pg' : 'essay'
    }

    return {
      pertanyaan,
      tipe_soal: finalTipe,
      opsi_a: finalTipe === 'pg' ? opsiA : null,
      opsi_b: finalTipe === 'pg' ? opsiB : null,
      opsi_c: finalTipe === 'pg' ? opsiC : null,
      opsi_d: finalTipe === 'pg' ? opsiD : null,
      jawaban_benar: jawaban || null,
      bobot: Number(bobot || 10),
      urutan: index + 1,
    }
  })
}

const validateQuestions = (questions) => {
  if (!questions.length) {
    throw new Error('Tidak ada soal yang berhasil dibaca dari file.')
  }

  questions.forEach((soal, index) => {
    if (!soal.pertanyaan) {
      throw new Error(`Soal nomor ${index + 1}: pertanyaan wajib diisi.`)
    }

    if (soal.tipe_soal === 'pg') {
      if (!soal.opsi_a || !soal.opsi_b || !soal.opsi_c || !soal.opsi_d) {
        throw new Error(
          `Soal nomor ${index + 1}: soal PG wajib punya opsi A, B, C, dan D.`,
        )
      }

      if (!soal.jawaban_benar) {
        throw new Error(`Soal nomor ${index + 1}: soal PG wajib punya kunci jawaban.`)
      }
    }
  })
}

const normalizeRole = (role) => {
  const value = String(role || '').toLowerCase().trim()

  if (value === 'guru') return 'guru'
  if (value === 'siswa') return 'siswa'
  if (value === 'staff') return 'admin'
  if (value === 'admin') return 'admin'

  return 'siswa'
}

const parseBoolean = (value) => {
  if (value === true) return true
  if (value === false) return false

  const text = String(value || '').toLowerCase().trim()
  return text === 'true' || text === '1' || text === 'aktif' || text === 'active'
}

const getValue = (source, keys, defaultValue = '') => {
  for (const key of keys) {
    if (source[key] !== undefined && source[key] !== null && source[key] !== '') {
      return source[key]
    }
  }

  return defaultValue
}

const parseCsvLine = (line) => {
  const result = []
  let current = ''
  let inQuotes = false

  for (let i = 0; i < line.length; i++) {
    const char = line[i]

    if (char === '"') {
      inQuotes = !inQuotes
    } else if (char === ',' && !inQuotes) {
      result.push(current.trim())
      current = ''
    } else {
      current += char
    }
  }

  result.push(current.trim())
  return result.map((item) => item.replace(/^"|"$/g, ''))
}

const parseCsvFile = (filePath) => {
  const content = fs.readFileSync(filePath, 'utf8')
  const lines = content
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)

  if (lines.length < 2) return []

  const headers = parseCsvLine(lines[0]).map((header) =>
    header.toLowerCase().trim(),
  )

  return lines.slice(1).map((line) => {
    const values = parseCsvLine(line)
    const row = {}

    headers.forEach((header, index) => {
      row[header] = values[index] || ''
    })

    return row
  })
}

const buildProfilePayload = ({ userId, body }) => {
  const nama = getValue(body, ['nama', 'name', 'Nama', 'Name'])
  const email = getValue(body, ['email', 'Email'])
  const role = normalizeRole(getValue(body, ['role', 'Role']))

  return {
    id: userId,
    nama,
    email,
    role,
    kelas: getValue(body, ['kelas', 'Kelas']) || null,
    is_active: body.is_active === undefined ? true : parseBoolean(body.is_active),

    nisn: getValue(body, ['nisn', 'NISN']) || null,
    nis: getValue(body, ['nis', 'NIS']) || null,
    nip: getValue(body, ['nip', 'NIP']) || null,
    nuptk: getValue(body, ['nuptk', 'NUPTK']) || null,

    tempat_lahir: getValue(body, ['tempat_lahir', 'tempat lahir']) || null,
    tanggal_lahir: getValue(body, ['tanggal_lahir', 'tanggal lahir']) || null,
    jenis_kelamin: getValue(body, ['jenis_kelamin', 'jenis kelamin']) || null,
    angkatan: getValue(body, ['angkatan', 'Angkatan']) || null,

    nama_ibu: getValue(body, ['nama_ibu', 'nama ibu']) || null,
    nama_ayah: getValue(body, ['nama_ayah', 'nama ayah']) || null,
    no_hp_ortu: getValue(body, ['no_hp_ortu', 'no hp ortu']) || null,
    pekerjaan_ortu: getValue(body, ['pekerjaan_ortu', 'pekerjaan ortu']) || null,

    status_kepegawaian:
      getValue(body, ['status_kepegawaian', 'status kepegawaian']) || null,
    pendidikan_terakhir:
      getValue(body, ['pendidikan_terakhir', 'pendidikan terakhir']) || null,
    mapel_utama: getValue(body, ['mapel_utama', 'mapel utama']) || null,
    wali_kelas: getValue(body, ['wali_kelas', 'wali kelas']) || null,
    jabatan_tambahan:
      getValue(body, ['jabatan_tambahan', 'jabatan tambahan']) || null,

    staff_id: getValue(body, ['staff_id', 'staff id']) || null,
    posisi_staff: getValue(body, ['posisi_staff', 'posisi staff']) || null,
  }
}

app.post('/api/admin/users', async (req, res) => {
  try {
    const { email, password } = req.body
    const nama = req.body.nama || req.body.name
    const role = normalizeRole(req.body.role)

    if (!nama || !email || !password || !role) {
      return res.status(400).json({
        success: false,
        error: 'Nama, email, password, dan role wajib diisi.',
      })
    }

    const { data: authData, error: authError } =
      await supabase.auth.admin.createUser({
        email,
        password,
        email_confirm: true,
        user_metadata: {
          nama,
          role,
        },
      })

    if (authError) throw authError

    const profilePayload = buildProfilePayload({
      userId: authData.user.id,
      body: {
        ...req.body,
        nama,
        email,
        role,
      },
    })

    const { error: profileError } = await supabase
      .from('profiles')
      .upsert(profilePayload, { onConflict: 'id' })

    if (profileError) throw profileError

    return res.json({
      success: true,
      message: 'User berhasil dibuat.',
      user: profilePayload,
    })
  } catch (err) {
    console.error(err)

    return res.status(500).json({
      success: false,
      error: 'Gagal membuat user.',
      detail: err.message,
    })
  }
})

app.post('/api/admin/users/import-csv', upload.single('file'), async (req, res) => {
  let uploadedFilePath = null

  try {
    if (!req.file) {
      return res.status(400).json({
        success: false,
        error: 'File CSV wajib diupload.',
      })
    }

    uploadedFilePath = req.file.path

    const rows = parseCsvFile(uploadedFilePath)
    const defaultRole = normalizeRole(req.body.default_role || 'siswa')

    const success = []
    const failed = []

    for (let i = 0; i < rows.length; i++) {
      try {
        const row = rows[i]

        const nama = getValue(row, ['nama', 'name'])
        const email = getValue(row, ['email'])
        const password = getValue(row, ['password'])
        const role = normalizeRole(getValue(row, ['role'], defaultRole))

        if (!nama || !email || !password) {
          throw new Error('nama, email, dan password wajib diisi.')
        }

        const { data: authData, error: authError } =
          await supabase.auth.admin.createUser({
            email,
            password,
            email_confirm: true,
            user_metadata: {
              nama,
              role,
            },
          })

        if (authError) throw authError

        const profilePayload = buildProfilePayload({
          userId: authData.user.id,
          body: {
            ...row,
            nama,
            email,
            role,
            is_active: true,
          },
        })

        const { error: profileError } = await supabase
          .from('profiles')
          .upsert(profilePayload, { onConflict: 'id' })

        if (profileError) throw profileError

        success.push({
          row: i + 2,
          email,
          nama,
          role,
        })
      } catch (err) {
        failed.push({
          row: i + 2,
          error: err.message,
        })
      }
    }

    return res.json({
      success: true,
      message: 'Import CSV selesai.',
      berhasil: success.length,
      gagal: failed.length,
      data_berhasil: success,
      data_gagal: failed,
    })
  } catch (err) {
    console.error(err)

    return res.status(500).json({
      success: false,
      error: 'Gagal import CSV user.',
      detail: err.message,
    })
  } finally {
    if (uploadedFilePath && fs.existsSync(uploadedFilePath)) {
      fs.unlinkSync(uploadedFilePath)
    }
  }
})

app.post('/api/kuis/import', upload.single('file'), async (req, res) => {
  let uploadedFilePath = null

  try {
    const { judul, kelas, deskripsi, kkm, guru_id, is_published, deadline } = req.body

    if (!req.file) {
      return res.status(400).json({ error: 'File tidak ada' })
    }

    uploadedFilePath = req.file.path

    if (!guru_id) {
      return res.status(400).json({ error: 'guru_id wajib dikirim.' })
    }

    if (!judul || !kelas || !kkm) {
      return res.status(400).json({
        error: 'Judul, kelas, dan KKM wajib diisi.',
      })
    }

    const ext = req.file.originalname.split('.').pop().toLowerCase()
    let soalData = []

    if (ext === 'xlsx' || ext === 'xls') {
      const rows = parseExcel(uploadedFilePath)
      soalData = parseExcelRowsToQuestions(rows)
    } else if (ext === 'docx') {
      const text = await parseDocx(uploadedFilePath)
      soalData = parseQuestionText(text)
    } else if (ext === 'pdf') {
      const text = await parsePdf(uploadedFilePath)
      soalData = parseQuestionText(text)
    } else {
      return res.status(400).json({
        error: 'Format file tidak didukung. Gunakan Excel, DOCX, atau PDF.',
      })
    }

    validateQuestions(soalData)

    const { data: kuis, error: kuisError } = await supabase
      .from('kuis')
      .insert({
        guru_id,
        judul,
        kelas,
        deskripsi: deskripsi || '',
        kkm: Number(kkm),
        mode_kuis: 'manual',
        is_published: is_published === 'true',
        deadline: deadline || null,
      })
      .select()
      .single()

    if (kuisError) throw kuisError

    const soalInsert = soalData.map((soal) => ({
      ...soal,
      kuis_id: kuis.id,
    }))

    const { error: soalError } = await supabase
      .from('kuis_soal')
      .insert(soalInsert)

    if (soalError) throw soalError

    return res.json({
      success: true,
      message: 'Kuis berhasil diimport',
      kuis_id: kuis.id,
      jumlah_soal: soalInsert.length,
      soal: soalInsert,
    })
  } catch (err) {
    console.error(err)

    return res.status(500).json({
      success: false,
      error: 'Gagal import kuis',
      detail: err.message,
    })
  } finally {
    if (uploadedFilePath && fs.existsSync(uploadedFilePath)) {
      fs.unlinkSync(uploadedFilePath)
    }
  }
})

app.listen(process.env.PORT || 3000, () => {
  console.log(`Server jalan di http://localhost:${process.env.PORT || 3000}`)
})