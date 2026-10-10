'use client'

import { useRef, useState } from 'react'
import Link from 'next/link'
import { useRouter } from 'next/navigation'
import * as XLSX from 'xlsx'
import { z } from 'zod'
import { ArrowLeft, CircleAlert, CircleCheck, Download, FileSpreadsheet, Upload } from 'lucide-react'
import { Badge } from '@/components/ui/badge'
import { Button } from '@/components/ui/button'
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table'
import { createQuestion } from '@/lib/question-store'
import { cn } from '@/lib/utils'

const rowSchema = z.object({
  question: z.string().trim().min(1, 'Question text is required'),
  optionA: z.string().min(1, 'Option A required'),
  optionB: z.string().min(1, 'Option B required'),
  optionC: z.string().optional().default(''),
  optionD: z.string().optional().default(''),
  correctAnswer: z.enum(['A', 'B', 'C', 'D'], { message: 'Correct answer must be A, B, C or D' }),
  category: z.enum(['traffic-signs', 'road-rules', 'defensive-driving', 'vehicle-knowledge'], { message: 'Unknown category' }),
  difficulty: z.enum(['easy', 'medium', 'hard'], { message: 'Difficulty must be easy, medium or hard' }),
  explanation: z.string().min(10, 'Explanation too short (min 10 chars)'),
  image: z.string().optional().default(''),
})

type ImportRow = z.infer<typeof rowSchema>

interface ParsedRow {
  row: number
  data: Record<string, unknown>
  result: { ok: true; value: ImportRow } | { ok: false; errors: string[] }
}

const TEMPLATE_COLUMNS = ['Question', 'Option A', 'Option B', 'Option C', 'Option D', 'Correct Answer', 'Category', 'Difficulty', 'Explanation', 'Image']

export default function ImportQuestionsPage() {
  const router = useRouter()
  const inputRef = useRef<HTMLInputElement>(null)
  const [fileName, setFileName] = useState<string | null>(null)
  const [rows, setRows] = useState<ParsedRow[] | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [imported, setImported] = useState(false)
  const [parsing, setParsing] = useState(false)

  const validCount = rows?.filter((r) => r.result.ok).length ?? 0
  const invalidCount = (rows?.length ?? 0) - validCount

  const handleFile = async (file: File | undefined) => {
    setError(null)
    setImported(false)
    if (!file) return
    const ext = file.name.split('.').pop()?.toLowerCase()
    if (!ext || !['xlsx', 'xls', 'csv'].includes(ext)) {
      setError('Unsupported file. Please upload an .xlsx, .xls or .csv spreadsheet.')
      return
    }
    setParsing(true)
    try {
      const buffer = await file.arrayBuffer()
      const workbook = XLSX.read(buffer, { type: 'array' })
      const sheet = workbook.Sheets[workbook.SheetNames[0]]
      if (!sheet) throw new Error('The spreadsheet has no sheets.')
      const json = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: '' })
      const normalizeHeader = (header: string) => header.replace(/^\uFEFF/, '').trim().toLowerCase().replace(/[\s_-]+/g, '')
      const getCell = (raw: Record<string, unknown>, expected: string) => {
        const key = Object.keys(raw).find((candidate) => normalizeHeader(candidate) === normalizeHeader(expected))
        return key ? raw[key] : ''
      }

      const parsed: ParsedRow[] = json.map((raw, i) => {
        const data = {
          question: String(getCell(raw, 'Question')).trim(),
          optionA: String(getCell(raw, 'Option A')).trim(),
          optionB: String(getCell(raw, 'Option B')).trim(),
          optionC: String(getCell(raw, 'Option C')).trim(),
          optionD: String(getCell(raw, 'Option D')).trim(),
          correctAnswer: String(getCell(raw, 'Correct Answer')).trim().toUpperCase(),
          category: String(getCell(raw, 'Category')).trim().toLowerCase(),
          difficulty: String(getCell(raw, 'Difficulty')).trim().toLowerCase(),
          explanation: String(getCell(raw, 'Explanation')).trim(),
          image: String(getCell(raw, 'Image')).trim(),
        }
        const result = rowSchema.safeParse(data)
        return result.success ? { row: i + 2, data, result: { ok: true, value: result.data } } : { row: i + 2, data, result: { ok: false, errors: result.error.issues.map((issue) => issue.message) } }
      })

      if (parsed.length === 0) {
        setError('No data rows found. Check that row 1 contains the column headers.')
        setRows(null)
      } else {
        setRows(parsed)
        setFileName(file.name)
      }
    } catch {
      setError('Could not parse this file. Make sure it is a valid Excel or CSV file using the template.')
    } finally {
      setParsing(false)
    }
  }

  const downloadTemplate = () => {
    const sample = [
      {
        Question: 'What does a flashing yellow traffic light mean?',
        'Option A': 'Stop and wait',
        'Option B': 'Proceed with caution',
        'Option C': 'Speed up',
        'Option D': 'Turn around',
        'Correct Answer': 'B',
        Category: 'traffic-signs',
        Difficulty: 'easy',
        Explanation: 'A flashing yellow light means slow down and proceed with caution.',
        Image: '',
      },
    ]
    const ws = XLSX.utils.json_to_sheet(sample, { header: TEMPLATE_COLUMNS })
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'Questions')
    XLSX.writeFile(wb, 'driveprep-question-template.xlsx')
  }

  const confirmImport = async () => {
    if (!rows) return
    setError(null)
    try {
      const validRows = rows.filter((r): r is ParsedRow & { result: { ok: true; value: ImportRow } } => r.result.ok)
      await Promise.all(validRows.map(({ result }) => {
        const v = result.value
        const options = [v.optionA, v.optionB, v.optionC, v.optionD].map((o) => o.trim()).filter(Boolean)
        const correctIndex = ['A', 'B', 'C', 'D'].indexOf(v.correctAnswer)
        return createQuestion({
          moduleId: v.category,
          type: 'single',
          text: v.question.trim(),
          options,
          correct: [correctIndex >= 0 ? correctIndex : 0],
          difficulty: v.difficulty,
          explanation: v.explanation.trim(),
        })
      }))
      setImported(true)
      router.push('/admin/questions')
      router.refresh()
    } catch {
      setError('Import failed. No questions were confirmed as imported. Please check the database connection and try again.')
    }
  }

  return (
    <div className="mx-auto max-w-5xl">
      <Link href="/admin/questions" className="inline-flex items-center gap-1.5 text-sm font-semibold text-muted-foreground hover:text-foreground">
        <ArrowLeft className="size-4" /> Questions
      </Link>
      <header className="mt-4 mb-6 flex flex-wrap items-center justify-between gap-3">
        <div>
          <h1 className="text-2xl font-extrabold tracking-tight sm:text-3xl">Import questions from Excel</h1>
          <p className="mt-1.5 text-sm text-muted-foreground sm:text-base">
            Parse → validate → preview → confirm. Invalid rows are reported before anything is committed.
          </p>
        </div>
        <Button variant="outline" className="h-10 gap-2 px-4 text-sm" onClick={downloadTemplate}>
          <Download className="size-4" /> Download template
        </Button>
      </header>

      <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={(e) => handleFile(e.target.files?.[0])} aria-label="Excel file" />

      {!rows && (
        <button
          type="button"
          onClick={() => inputRef.current?.click()}
          className="flex w-full flex-col items-center justify-center gap-3 rounded-2xl border-2 border-dashed border-border bg-card px-6 py-16 text-center transition-colors hover:border-primary/50 hover:bg-primary/5"
        >
          <span className="flex size-14 items-center justify-center rounded-full bg-primary/10">
            <FileSpreadsheet className="size-6 text-primary" />
          </span>
          <span className="text-sm font-semibold">{parsing ? 'Parsing spreadsheet…' : 'Click to select your .xlsx file'}</span>
          <span className="text-xs text-muted-foreground">Columns: {TEMPLATE_COLUMNS.join(' · ')}</span>
        </button>
      )}

      {error && (
        <p className="mt-4 flex items-center gap-2 rounded-lg bg-destructive/10 px-3.5 py-2.5 text-sm font-medium text-destructive">
          <CircleAlert className="size-4 shrink-0" /> {error}
        </p>
      )}

      {rows && (
        <>
          <section className="flex flex-wrap items-center gap-3 rounded-2xl border border-border/70 bg-card p-5 shadow-[0_1px_3px_rgb(16_24_40/0.05)]">
            <FileSpreadsheet className="size-5 text-green-600" />
            <div className="min-w-0 flex-1">
              <p className="truncate text-sm font-bold">{fileName}</p>
              <p className="text-xs text-muted-foreground">
                {rows.length} rows parsed · <span className="font-semibold text-green-600">{validCount} valid</span>
                {invalidCount > 0 && <> · <span className="font-semibold text-destructive">{invalidCount} invalid</span></>}
              </p>
            </div>
            <div className="flex gap-2.5">
              <Button variant="outline" className="h-10 px-4 text-sm" onClick={() => { setRows(null); setFileName(null); if (inputRef.current) inputRef.current.value = '' }}>
                Choose another file
              </Button>
              <Button className="h-10 gap-2 px-5 text-sm" disabled={validCount === 0 || imported} onClick={confirmImport}>
                {imported ? (
                  <>
                    <CircleCheck className="size-4" /> Imported
                  </>
                ) : (
                  <>
                    <Upload className="size-4" /> Confirm import ({validCount})
                  </>
                )}
              </Button>
            </div>
          </section>

          {imported && (
            <p className="mt-4 flex items-center gap-2 rounded-lg bg-green-50 px-3.5 py-2.5 text-sm font-medium text-green-700 ring-1 ring-green-100">
              <CircleCheck className="size-4" /> {validCount} questions imported successfully. They are now part of the question bank.
            </p>
          )}

          <section className="mt-4 overflow-hidden rounded-2xl border border-border/70 bg-card shadow-[0_1px_3px_rgb(16_24_40/0.05)]">
            <div className="border-b border-border/70 px-5 py-3.5">
              <h2 className="text-sm font-extrabold">Preview — nothing is inserted until you confirm</h2>
            </div>
            <Table>
              <TableHeader>
                <TableRow>
                  <TableHead className="pl-5">Row</TableHead>
                  <TableHead>Question</TableHead>
                  <TableHead>Category</TableHead>
                  <TableHead>Difficulty</TableHead>
                  <TableHead className="pr-5">Validation</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {rows.map((r) => (
                  <TableRow key={r.row} className={cn(!r.result.ok && 'bg-red-50/50')}>
                    <TableCell className="pl-5 text-muted-foreground">{r.row}</TableCell>
                    <TableCell className="max-w-sm">
                      <p className="truncate font-semibold">{String(r.data.question) || '—'}</p>
                    </TableCell>
                    <TableCell className="text-muted-foreground">{String(r.data.category) || '—'}</TableCell>
                    <TableCell className="text-muted-foreground">{String(r.data.difficulty) || '—'}</TableCell>
                    <TableCell className="pr-5">
                      {r.result.ok ? (
                        <Badge variant="success">Valid</Badge>
                      ) : (
                        <div className="flex flex-col items-start gap-1">
                          <Badge variant="destructive">Invalid</Badge>
                          {r.result.errors.slice(0, 2).map((err, i) => (
                            <p key={i} className="text-xs text-destructive">
                              {err}
                            </p>
                          ))}
                        </div>
                      )}
                    </TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </section>
        </>
      )}
    </div>
  )
}
