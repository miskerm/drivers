'use client'

import { useRef, useState } from 'react'
import Link from 'next/link'
import * as XLSX from 'xlsx'
import mammoth from 'mammoth'
import { z } from 'zod'
import { ArrowLeft, CircleAlert, CircleCheck, Download, FileSpreadsheet, Upload } from 'lucide-react'
import { Badge } from '@/components/ui/badge'
import { Button } from '@/components/ui/button'
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table'
import type { Question } from '@/lib/types'

const questionSchema = z.object({
  question: z.string().trim().min(1), optionA: z.string().trim().optional().default(''), optionB: z.string().trim().optional().default(''),
  optionC: z.string().trim().optional().default(''), optionD: z.string().trim().optional().default(''), category: z.string().trim().optional().default('road-rules'),
  difficulty: z.string().trim().optional().default('easy'), image: z.string().optional().default(''),
})
const answerSchema = z.object({ answer: z.string().trim().min(1) })
type QuestionRow = z.infer<typeof questionSchema> & { number: number; sourceRow: number }
type AnswerRow = z.infer<typeof answerSchema> & { number: number; explanation: string; sourceRow: number }
type PreviewRow = { number: number; question: QuestionRow; answer: AnswerRow; valid: boolean; errors: string[] }

const columns = ['Question Number', 'Question', 'Option A', 'Option B', 'Option C', 'Option D', 'Category', 'Difficulty', 'Image']
const clean = (value: unknown) => String(value ?? '').replace(/\s+/g, ' ').trim()
const key = (value: unknown) => clean(value).toLowerCase().replace(/[^a-z0-9]/g, '')
const labels = {
  number: new Set(['number', 'no', 'id', 'questionnumber', 'questionno']), question: new Set(['question', 'questiontext', 'prompt']),
  answer: new Set(['answer', 'rightanswer', 'correctanswer', 'correct', 'key', 'right']), explanation: new Set(['explanation', 'instruction', 'rationale', 'reason']),
  option: new Set(['optiona', 'optionb', 'optionc', 'optiond', 'choicea', 'choiceb', 'choicec', 'choiced', 'a', 'b', 'c', 'd']),
  category: new Set(['category', 'module']), difficulty: new Set(['difficulty', 'level']), image: new Set(['image', 'imageurl']),
}
const allLabels = new Set(Object.values(labels).flatMap((set) => [...set]))

function sheetRows(file: ArrayBuffer) {
  const workbook = XLSX.read(file, { type: 'array' })
  return workbook.SheetNames.flatMap((name) => XLSX.utils.sheet_to_json<unknown[]>(workbook.Sheets[name], { header: 1, defval: '' }))
}
function sequence(numbers: number[]) {
  const sorted = [...numbers].sort((a, b) => a - b)
  return sorted.length > 0 && sorted.every((number, index) => number === index + 1)
}
function parseDocumentText(text: string) {
  return text.split(/\r?\n/).map((line) => [line]).filter((row) => clean(row[0]))
}
async function documentRows(file: File) {
  const buffer = await file.arrayBuffer()
  if (file.name.toLowerCase().endsWith('.docx')) {
    const result = await mammoth.extractRawText({ arrayBuffer: buffer })
    return parseDocumentText(result.value)
  }
  const workbook = XLSX.read(buffer, { type: 'array' })
  return workbook.SheetNames.flatMap((name) => XLSX.utils.sheet_to_json<unknown[]>(workbook.Sheets[name], { header: 1, defval: '' }))
}
function readQuestionFile(rows: unknown[][]): QuestionRow[] {
  const parsed: QuestionRow[] = []
  let current: Partial<QuestionRow> | null = null

  const finish = (sourceRow: number) => {
    if (!current?.number || !current.question) return
    const result = questionSchema.safeParse(current)
    if (result.success && !/^write or paste question/i.test(result.data.question)) {
      parsed.push({ ...result.data, number: current.number, sourceRow })
    }
  }

  rows.forEach((cells, rowIndex) => {
    const values = cells.map(clean).filter(Boolean)
    if (!values.length) return
    const text = values.join(' ')
    // The supplied 50-question template stores each question as “Q1 text” and
    // its choices as “A) text”, “B) text”, etc., without column headers.
    const questionMatch = text.match(/^(?:Q\s*)?(\d+)\s*(?:[:.)-]|\s+)\s*(.*)$/i)
    if (questionMatch) {
      finish(rowIndex)
      current = { number: Number(questionMatch[1].replace(/\s/g, '')), question: clean(questionMatch[2]) }
      return
    }
    const optionMatch = text.match(/^([A-D])\s*[).:-]\s*(.*)$/i)
    if (optionMatch && current) {
      current[`option${optionMatch[1].toUpperCase()}` as 'optionA' | 'optionB' | 'optionC' | 'optionD'] = clean(optionMatch[2])
      return
    }

    const data: Record<string, string> = {}; let number = 0
    cells.forEach((cell, i) => { const value = clean(cell); const label = key(value); const next = clean(cells[i + 1])
      if (labels.number.has(label) && /^\d+$/.test(next)) number = Number(next)
      else if (labels.question.has(label)) data.question = next
      else if (labels.option.has(label)) data[`option${label.slice(-1).toUpperCase()}`] = next
      else if (labels.category.has(label)) data.category = next
      else if (labels.difficulty.has(label)) data.difficulty = next
      else if (labels.image.has(label)) data.image = next
    })
    if (!data.question) data.question = values.find((value) => !allLabels.has(key(value)) && !/^\d+$/.test(value)) ?? ''
    if (!number) number = Number(values.find((value) => /^\d+$/.test(value)) ?? 0)
    const result = questionSchema.safeParse({ ...data, question: data.question })
    if (result.success && number > 0) parsed.push({ ...result.data, number, sourceRow: rowIndex + 1 })
  })
  finish(rows.length)
  return parsed
}
function readAnswerFile(rows: unknown[][]): AnswerRow[] {
  return rows.flatMap((cells, rowIndex) => {
    const values = cells.map(clean).filter(Boolean); if (!values.length) return []
    let number = Number(values.find((value) => /^\d+$/.test(value)) ?? 0); let answer = ''; let explanation = ''
    cells.forEach((cell, i) => { const label = key(cell); const next = clean(cells[i + 1]);
      if (labels.number.has(label) && /^\d+$/.test(next)) number = Number(next)
      else if (labels.answer.has(label)) answer = next
      else if (labels.explanation.has(label)) explanation = next
    })
    if (!answer) answer = values.find((value) => !/^\d+$/.test(value) && !allLabels.has(key(value))) ?? ''
    const result = answerSchema.safeParse({ answer })
    return result.success && number > 0 ? [{ ...result.data, number, explanation, sourceRow: rowIndex + 1 }] : []
  })
}

export default function ImportQuestionsPage() {
  const questionInput = useRef<HTMLInputElement>(null), answerInput = useRef<HTMLInputElement>(null)
  const [questionFile, setQuestionFile] = useState<string | null>(null), [answerFile, setAnswerFile] = useState<string | null>(null)
  const [questions, setQuestions] = useState<QuestionRow[] | null>(null), [answers, setAnswers] = useState<AnswerRow[] | null>(null)
  const [preview, setPreview] = useState<PreviewRow[] | null>(null), [error, setError] = useState<string | null>(null), [imported, setImported] = useState(false)
  const readFile = async (file: File, kind: 'question' | 'answer') => {
    setError(null); setImported(false)
    try { const rows = await documentRows(file); const parsed = kind === 'question' ? readQuestionFile(rows) : readAnswerFile(rows)
      if (!parsed.length) throw new Error(`No ${kind}s were detected. Include a number starting at 1 and the ${kind} text.`)
      if (!sequence(parsed.map((row) => row.number))) throw new Error(`${kind === 'question' ? 'Questions' : 'Answers'} must be numbered consecutively starting at 1.`)
      if (kind === 'question') { setQuestions(parsed as QuestionRow[]); setQuestionFile(file.name) } else { setAnswers(parsed as AnswerRow[]); setAnswerFile(file.name) }
      setPreview(null)
    } catch (e) { setError(e instanceof Error ? e.message : 'Could not read this Excel file.') }
  }
  const matchFiles = () => {
    if (!questions || !answers) return
    const answerMap = new Map(answers.map((answer) => [answer.number, answer])); setPreview(questions.map((question) => {
      const answer = answerMap.get(question.number); return { number: question.number, question, answer: answer!, valid: !!answer, errors: answer ? [] : ['No answer with this number was found.'] }
    }))
  }
  const confirmImport = async () => {
    if (!preview) return
    const payload: Omit<Question, 'id'>[] = preview.filter((row) => row.valid).map(({ question, answer }) => {
      const options = [question.optionA, question.optionB, question.optionC, question.optionD].map(clean).filter(Boolean); if (!options.length) options.push(answer.answer)
      const raw = answer.answer.toLowerCase(); const index = ['a', 'b', 'c', 'd'].indexOf(raw) >= 0 ? ['a', 'b', 'c', 'd'].indexOf(raw) : ['1', '2', '3', '4'].indexOf(raw) >= 0 ? Number(raw) - 1 : Math.max(0, options.findIndex((option) => option.toLowerCase() === raw))
      return { moduleId: question.category, type: 'single', text: question.question, options, correct: [index], difficulty: ['easy', 'medium', 'hard'].includes(question.difficulty.toLowerCase()) ? question.difficulty.toLowerCase() as Question['difficulty'] : 'easy', explanation: answer.explanation || 'Review the rule and select the best answer.' }
    })
    try { const response = await fetch('/api/admin/questions', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ questions: payload }) }); if (!response.ok) throw new Error(); setImported(true) } catch { setError('The files were read, but the questions could not be saved to Turso.') }
  }
  const downloadTemplate = () => { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['Instructions'], ['Upload Questions first, numbered 1, 2, 3…'], ['Then upload Answers, using the same numbers.'], ['Answer sheet may include Explanation.']]), 'Instructions'); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([{ 'Question Number': 1, Question: 'Write question here', 'Option A': '', 'Option B': '', 'Option C': '', 'Option D': '', Category: 'road-rules', Difficulty: 'easy', Image: '' }], { header: columns }), 'Questions'); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([{ 'Question Number': 1, 'Right Answer': 'B', Explanation: 'Optional explanation' }]), 'Answers'); XLSX.writeFile(wb, 'driveprep-question-answer-template.xlsx') }
  const valid = preview?.filter((row) => row.valid).length ?? 0
  return <div className="mx-auto max-w-5xl"><Link href="/admin/questions" className="inline-flex items-center gap-1.5 text-sm font-semibold text-muted-foreground"><ArrowLeft className="size-4" /> Questions</Link><header className="mt-4 mb-6 flex flex-wrap items-center justify-between gap-3"><div><h1 className="text-2xl font-extrabold sm:text-3xl">Import questions from Excel</h1><p className="mt-1.5 text-sm text-muted-foreground">Upload questions first as .xlsx, .csv, or .docx. After the numbered questions are valid, upload the separate numbered answers file; matching uses numbers only.</p></div><Button variant="outline" onClick={downloadTemplate}><Download className="size-4" /> Download template</Button></header>
    <div className="grid gap-4 md:grid-cols-2"><UploadCard title="1. Questions file" fileName={questionFile} input={questionInput} onFile={(file) => file && readFile(file, 'question')} /><UploadCard title="2. Answers file" fileName={answerFile} disabled={!questions} input={answerInput} onFile={(file) => file && readFile(file, 'answer')} /></div>
    {error && <p className="mt-4 flex items-center gap-2 rounded-lg bg-destructive/10 px-3.5 py-2.5 text-sm font-medium text-destructive"><CircleAlert className="size-4" /> {error}</p>}
    {questions && answers && <div className="mt-4 flex items-center gap-3 rounded-2xl border bg-card p-5"><FileSpreadsheet className="size-5 text-green-600" /><p className="flex-1 text-sm font-semibold">{questions.length} questions and {answers.length} answers loaded{preview ? ` · ${valid} matched` : ''}</p><Button onClick={matchFiles}>Validate and match</Button>{preview && <Button disabled={!valid || imported} onClick={confirmImport}>{imported ? <><CircleCheck className="size-4" /> Imported</> : <><Upload className="size-4" /> Import {valid}</>}</Button>}</div>}
    {preview && <section className="mt-4 overflow-hidden rounded-2xl border bg-card"><div className="border-b px-5 py-3.5"><h2 className="text-sm font-extrabold">Preview — nothing is inserted until import</h2></div><Table><TableHeader><TableRow><TableHead className="pl-5">Number</TableHead><TableHead>Question</TableHead><TableHead>Answer</TableHead><TableHead>Validation</TableHead></TableRow></TableHeader><TableBody>{preview.map((row) => <TableRow key={row.number}><TableCell className="pl-5">{row.number}</TableCell><TableCell className="max-w-sm truncate">{row.question.question}</TableCell><TableCell>{row.answer?.answer || '—'}</TableCell><TableCell>{row.valid ? <Badge variant="success">Question and answer valid</Badge> : <Badge variant="destructive">Answer missing</Badge>}</TableCell></TableRow>)}</TableBody></Table></section>}
  </div>
}
function UploadCard({ title, fileName, disabled = false, input, onFile }: { title: string; fileName: string | null; disabled?: boolean; input: React.RefObject<HTMLInputElement | null>; onFile: (file?: File) => void }) { return <div className="rounded-2xl border-2 border-dashed bg-card p-6 text-center"><input ref={input} type="file" accept=".xlsx,.xls,.csv,.docx" className="hidden" disabled={disabled} onChange={(event) => onFile(event.target.files?.[0])} /><Button variant="outline" disabled={disabled} onClick={() => input.current?.click()}><FileSpreadsheet className="size-4" /> {fileName || title}</Button>{disabled && <p className="mt-2 text-xs text-muted-foreground">Upload and validate the questions first</p>}{fileName && <p className="mt-2 text-xs text-muted-foreground">Loaded and ready to validate</p>}</div> }

/* The importer deliberately treats spreadsheet coordinates as optional; only the detected number and content matter. */
// eslint-disable-next-line @typescript-eslint/no-unused-vars
const _unused = Upload
                                                                    
