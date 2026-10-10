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
import type { Question } from '@/lib/types'
import { cn } from '@/lib/utils'

const questionSchema = z.object({
  question: z.string().trim().min(1, 'Question is required'),
  optionA: z.string().trim().optional().default(''),
  optionB: z.string().trim().optional().default(''),
  optionC: z.string().trim().optional().default(''),
  optionD: z.string().trim().optional().default(''),
  category: z.string().trim().optional().default('road-rules'),
  difficulty: z.string().trim().optional().default('easy'),
  explanation: z.string().optional().default(''),
  image: z.string().optional().default(''),
})

const answerSchema = z.object({ correctAnswer: z.string().trim().min(1, 'Right answer is required') })
type ImportRow = z.infer<typeof questionSchema> & z.infer<typeof answerSchema>

interface ParsedRow {
  row: number
  data: Record<string, unknown>
  questionResult: { ok: true; value: z.infer<typeof questionSchema> } | { ok: false; errors: string[] }
  answerResult: { ok: true; value: string } | { ok: false; errors: string[] }
}

const TEMPLATE_COLUMNS = ['Question Number', 'Question', 'Option A', 'Option B', 'Option C', 'Option D', 'Right Answer', 'Category', 'Difficulty', 'Instruction / Explanation', 'Image']

export default function ImportQuestionsPage() {
  const router = useRouter()
  const inputRef = useRef<HTMLInputElement>(null)
  const [fileName, setFileName] = useState<string | null>(null)
  const [rows, setRows] = useState<ParsedRow[] | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [imported, setImported] = useState(false)
  const [parsing, setParsing] = useState(false)

  const validCount = rows?.filter((r) => r.questionResult.ok && r.answerResult.ok).length ?? 0
  const questionCount = rows?.filter((r) => r.questionResult.ok).length ?? 0
  const answerCount = rows?.filter((r) => r.answerResult.ok).length ?? 0
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
      const key = (value: string) => value.toLowerCase().replace(/[^a-z0-9]/g, '')
      const clean = (value: unknown) => String(value ?? '').replace(/\s+/g, ' ').trim()
      const questionLabels = new Set(['question', 'questiontext', 'prompt', 'questiontitle'])
      const answerLabels = new Set(['answer', 'rightanswer', 'correctanswer', 'correct', 'key', 'right'])
      const optionLabels = new Set(['optiona', 'optionb', 'optionc', 'optiond', 'choicea', 'choiceb', 'choicec', 'choiced', 'a', 'b', 'c', 'd'])
      const metadataLabels = new Set(['category', 'module', 'difficulty', 'level', 'explanation', 'rationale', 'reason', 'instruction', 'image', 'imageurl'])

      // Read every cell as text. Headers, row numbers, and column positions are optional.
      // A question can be followed by its answer in the same row or in the next cell/row.
      const candidateSheets = workbook.SheetNames.map((name) => ({
        name,
        rows: XLSX.utils.sheet_to_json<unknown[]>(workbook.Sheets[name], { header: 1, defval: '' }),
      }))
      const source = candidateSheets.find(({ rows }) => rows.some((cells) => cells.some((cell) => questionLabels.has(key(clean(cell))))) ) ?? candidateSheets[0]
      if (!source) throw new Error('The spreadsheet has no sheets.')

      const parsed: ParsedRow[] = []
      let pendingQuestion = ''
      source.rows.forEach((cells, index) => {
        const values = cells.map(clean).filter(Boolean)
        if (!values.length) return
        const data = {
          question: '', optionA: '', optionB: '', optionC: '', optionD: '', correctAnswer: '',
          category: 'road-rules', difficulty: 'easy', explanation: '', image: '',
        }
        let unlabeled: string[] = []
        cells.forEach((cell, cellIndex) => {
          const value = clean(cell)
          if (!value) return
          const label = key(value)
          const next = clean(cells[cellIndex + 1])
          if (questionLabels.has(label)) data.question = next
          else if (answerLabels.has(label)) data.correctAnswer = next
          else if (optionLabels.has(label)) data[`option${label.slice(-1).toUpperCase()}` as 'optionA' | 'optionB' | 'optionC' | 'optionD'] = next
          else if (metadataLabels.has(label)) {
            const field = label === 'module' ? 'category' : label === 'level' ? 'difficulty' : label === 'rationale' || label === 'reason' || label === 'instruction' ? 'explanation' : label === 'imageurl' ? 'image' : label as keyof typeof data
            if (field in data) data[field] = next
          } else if (!Object.values([...questionLabels, ...answerLabels, ...optionLabels, ...metadataLabels]).includes(label)) unlabeled.push(value)
        })
        if (!data.question && unlabeled.length) data.question = pendingQuestion || unlabeled[0]
        if (!data.correctAnswer && unlabeled.length > (pendingQuestion ? 0 : 1)) data.correctAnswer = unlabeled[pendingQuestion ? 0 : 1]
        if (data.question && data.correctAnswer) {
          pendingQuestion = ''
          const question = questionSchema.safeParse(data)
          const answer = answerSchema.safeParse({ correctAnswer: data.correctAnswer })
          parsed.push({ row: index + 1, data, questionResult: question.success ? { ok: true, value: question.data } : { ok: false, errors: question.error.issues.map((issue) => issue.message) }, answerResult: answer.success ? { ok: true, value: answer.data.correctAnswer } : { ok: false, errors: answer.error.issues.map((issue) => issue.message) } })
        } else if (data.question) pendingQuestion = data.question
      })

      if (parsed.length === 0) {
        setError('No question-and-answer pairs found. The importer ignores row and column numbers; put a question and its answer anywhere in the sheet.')
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
        'Question Number': 1,
        Question: 'What does a flashing yellow traffic light mean?',
        'Option A': 'Stop and wait',
        'Option B': 'Proceed with caution',
        'Option C': 'Speed up',
        'Option D': 'Turn around',
        'Right Answer': 'B',
        Category: 'traffic-signs',
        Difficulty: 'easy',
        'Instruction / Explanation': 'A flashing yellow light means slow down and proceed with caution.',
        Image: '',
      },
      {
        'Question Number': 2,
        Question: 'Write the next question here.',
        'Option A': '',
        'Option B': '',
        'Option C': '',
        'Option D': '',
        'Right Answer': 'Write the correct answer here.',
        Category: 'road-rules',
        Difficulty: 'easy',
        'Instruction / Explanation': 'Optional instructions or explanation.',
        Image: '',
      },
    ]
    const ws = XLSX.utils.json_to_sheet(sample, { header: TEMPLATE_COLUMNS })
    const instructions = [
      ['DrivePrep question import template'],
      ['Instructions'],
      ['1. Keep the header names in row 1. Add questions starting at row 2 and continue numbering from 1 upward; any number of rows is supported.'],
      ['2. Question and Right Answer are required. Option A-D are optional for a question-and-answer format.'],
      ['3. Right Answer may be A, B, C, D, 1, 2, 3, 4, or the exact answer text.'],
      ['4. Category, Difficulty, Image, and Instruction / Explanation are optional.'],
      ['5. Downloaded rows are validated first. Only valid rows are saved to the Turso question bank after confirmation.'],
    ]
    const instructionSheet = XLSX.utils.aoa_to_sheet(instructions)
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, instructionSheet, 'Instructions')
    XLSX.utils.book_append_sheet(wb, ws, 'Questions')
    XLSX.writeFile(wb, 'driveprep-question-template.xlsx')
  }

  const confirmImport = async () => {
    if (!rows) return
    setError(null)
    const questions: Omit<Question, 'id'>[] = rows.flatMap((r) => {
      if (!r.questionResult.ok || !r.answerResult.ok) return []
      const v = { ...r.questionResult.value, correctAnswer: r.answerResult.value }
      const answerText = v.correctAnswer.trim()
      const options = [v.optionA, v.optionB, v.optionC, v.optionD].map((o) => o.trim()).filter(Boolean)
      const answer = answerText.toLowerCase()
      if (options.length === 0) options.push(answerText)
      const letterIndex = ['a', 'b', 'c', 'd'].indexOf(answer.replace(/[^a-d]/g, ''))
      const numericIndex = ['1', '2', '3', '4'].indexOf(answer)
      const textIndex = options.findIndex((option) => option.toLowerCase() === answer)
      const correctIndex = letterIndex >= 0 ? letterIndex : numericIndex >= 0 ? numericIndex : textIndex
      const difficulty = ['easy', 'medium', 'hard'].includes(v.difficulty.toLowerCase()) ? v.difficulty.toLowerCase() as Question['difficulty'] : 'easy'
      return [{ moduleId: v.category, type: 'single', text: v.question.trim(), options, correct: [correctIndex >= 0 ? correctIndex : 0], difficulty, explanation: v.explanation.trim() || 'Review the rule and select the best answer.' }]
    })
    try {
      const response = await fetch('/api/admin/questions', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ questions }) })
      if (!response.ok) throw new Error('Import failed')
      setImported(true)
    } catch {
      setError('The spreadsheet was read, but the questions could not be saved to the database. Please try again.')
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
                {rows.length} rows parsed · <span className="font-semibold text-green-600">{questionCount} questions valid</span> · <span className="font-semibold text-green-600">{answerCount} answers valid</span>
                {invalidCount > 0 && <> · <span className="font-semibold text-destructive">{invalidCount} incomplete</span></>}
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
                  <TableRow key={r.row} className={cn(!(r.questionResult.ok && r.answerResult.ok) && 'bg-red-50/50')}>
                    <TableCell className="pl-5 text-muted-foreground">{r.row}</TableCell>
                    <TableCell className="max-w-sm">
                      <p className="truncate font-semibold">{String(r.data.question) || '—'}</p>
                    </TableCell>
                    <TableCell className="text-muted-foreground">{String(r.data.category) || '—'}</TableCell>
                    <TableCell className="text-muted-foreground">{String(r.data.difficulty) || '—'}</TableCell>
                    <TableCell className="pr-5">
                      <div className="flex flex-col items-start gap-1">
                        <div className="flex gap-1.5">
                          {r.questionResult.ok ? <Badge variant="success">Question valid</Badge> : <Badge variant="destructive">Question invalid</Badge>}
                          {r.answerResult.ok ? <Badge variant="success">Answer valid</Badge> : <Badge variant="destructive">Answer invalid</Badge>}
                        </div>
                        {!r.questionResult.ok && r.questionResult.errors.slice(0, 1).map((err, i) => <p key={`q-${i}`} className="text-xs text-destructive">Question: {err}</p>)}
                        {!r.answerResult.ok && r.answerResult.errors.slice(0, 1).map((err, i) => <p key={`a-${i}`} className="text-xs text-destructive">Answer: {err}</p>)}
                      </div>
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
