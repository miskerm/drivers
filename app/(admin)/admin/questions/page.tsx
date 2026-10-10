'use client'

import { useMemo, useState } from 'react'
import Link from 'next/link'
import * as XLSX from 'xlsx'
import { Download, Pencil, Plus, Search, Trash2, Upload } from 'lucide-react'
import { Badge } from '@/components/ui/badge'
import { ConfirmDialog } from '@/components/ui/confirm-dialog'
import { Input } from '@/components/ui/input'
import { Select } from '@/components/ui/select'
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table'
import { useModules } from '@/lib/module-store'
import { deleteQuestion, useQuestions, type Question } from '@/lib/question-store'

export default function AdminQuestionsPage() {
  const questions = useQuestions()
  const { modules } = useModules()
  const [query, setQuery] = useState('')
  const [moduleFilter, setModuleFilter] = useState('all')
  const [difficultyFilter, setDifficultyFilter] = useState('all')
  const [pendingDelete, setPendingDelete] = useState<Question | null>(null)
  const [deleted, setDeleted] = useState(false)

  const moduleTitle = useMemo(() => Object.fromEntries(modules.map((m) => [m.id, m.title])), [modules])

  const filtered = useMemo(
    () =>
      questions.filter((q) => {
        const matchesQuery = q.text.toLowerCase().includes(query.toLowerCase())
        const matchesModule = moduleFilter === 'all' || q.moduleId === moduleFilter
        const matchesDifficulty = difficultyFilter === 'all' || q.difficulty === difficultyFilter
        return matchesQuery && matchesModule && matchesDifficulty
      }),
    [questions, query, moduleFilter, difficultyFilter],
  )

  const confirmDelete = () => {
    if (!pendingDelete) return
    deleteQuestion(pendingDelete.id)
    setPendingDelete(null)
    setDeleted(true)
    setTimeout(() => setDeleted(false), 4000)
  }

  const exportQuestions = () => {
    const data = filtered.map((question) => ({
      Question: question.text,
      'Option A': question.options[0] ?? '',
      'Option B': question.options[1] ?? '',
      'Option C': question.options[2] ?? '',
      'Option D': question.options[3] ?? '',
      'Correct Answer': question.correct.map((index) => ['A', 'B', 'C', 'D'][index] ?? '').join(', '),
      Category: question.moduleId,
      Difficulty: question.difficulty,
      Explanation: question.explanation,
    }))
    const worksheet = XLSX.utils.json_to_sheet(data)
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Questions')
    XLSX.writeFile(workbook, 'driveprep-questions.xlsx')
  }

  return (
    <div className="mx-auto max-w-6xl">
      <header className="flex flex-wrap items-center justify-between gap-4">
        <div>
          <h1 className="text-2xl font-extrabold tracking-tight sm:text-3xl">Questions</h1>
          <p className="mt-1.5 text-sm text-muted-foreground sm:text-base">{questions.length} questions in the bank across all modules.</p>
        </div>
        <div className="flex gap-2.5">
          <button type="button" onClick={exportQuestions} className="flex h-10 items-center gap-2 rounded-lg border border-border bg-card px-4 text-sm font-semibold transition-colors hover:bg-muted">
            <Download className="size-4" /> Export Excel
          </button>
          <Link href="/admin/questions/import" className="flex h-10 items-center gap-2 rounded-lg border border-border bg-card px-4 text-sm font-semibold transition-colors hover:bg-muted">
            <Upload className="size-4" /> Import Excel
          </Link>
          <Link href="/admin/questions/new" className="flex h-10 items-center gap-2 rounded-lg bg-primary px-4 text-sm font-semibold text-primary-foreground transition-colors hover:bg-primary/90">
            <Plus className="size-4" /> New question
          </Link>
        </div>
      </header>

      {deleted && (
        <p className="mt-4 rounded-lg bg-green-50 px-3.5 py-2.5 text-sm font-medium text-green-700 ring-1 ring-green-100">
          Question deleted.
        </p>
      )}

      <div className="mt-6 grid gap-3 sm:grid-cols-[1.5fr_1fr_1fr]">
        <div className="relative">
          <Search className="absolute top-1/2 left-3 size-4 -translate-y-1/2 text-muted-foreground" />
          <Input placeholder="Search questions…" className="pl-9" value={query} onChange={(e) => setQuery(e.target.value)} />
        </div>
        <Select value={moduleFilter} onChange={(e) => setModuleFilter(e.target.value)} aria-label="Filter by module">
          <option value="all">All modules</option>
          {modules.map((m) => (
            <option key={m.id} value={m.id}>
              {m.title}
            </option>
          ))}
        </Select>
        <Select value={difficultyFilter} onChange={(e) => setDifficultyFilter(e.target.value)} aria-label="Filter by difficulty">
          <option value="all">All difficulties</option>
          <option value="easy">Easy</option>
          <option value="medium">Medium</option>
          <option value="hard">Hard</option>
        </Select>
      </div>

      <section className="mt-4 rounded-2xl border border-border/70 bg-card shadow-[0_1px_3px_rgb(16_24_40/0.05)]">
        <Table>
          <TableHeader>
            <TableRow>
              <TableHead className="pl-5">Question</TableHead>
              <TableHead>Module</TableHead>
              <TableHead>Type</TableHead>
              <TableHead>Difficulty</TableHead>
              <TableHead className="pr-5 text-right">Actions</TableHead>
            </TableRow>
          </TableHeader>
          <TableBody>
            {filtered.map((q) => (
              <TableRow key={q.id}>
                <TableCell className="max-w-md pl-5">
                  <p className="truncate font-semibold">{q.text}</p>
                  <p className="truncate text-xs text-muted-foreground">{q.explanation}</p>
                </TableCell>
                <TableCell>
                  <Badge variant="secondary">{moduleTitle[q.moduleId] ?? q.moduleId}</Badge>
                </TableCell>
                <TableCell className="text-muted-foreground">{q.type === 'truefalse' ? 'True/False' : 'Choice'}</TableCell>
                <TableCell className="capitalize">{q.difficulty}</TableCell>
                <TableCell className="pr-5">
                  <div className="flex items-center justify-end gap-1.5">
                    <Link
                      href={`/admin/questions/${q.id}/edit`}
                      className="flex size-8 items-center justify-center rounded-lg text-muted-foreground transition-colors hover:bg-muted hover:text-foreground"
                      aria-label={`Edit question ${q.id}`}
                    >
                      <Pencil className="size-4" />
                    </Link>
                    <button
                      onClick={() => setPendingDelete(q)}
                      className="flex size-8 items-center justify-center rounded-lg text-muted-foreground transition-colors hover:bg-red-50 hover:text-destructive"
                      aria-label={`Delete question ${q.id}`}
                    >
                      <Trash2 className="size-4" />
                    </button>
                  </div>
                </TableCell>
              </TableRow>
            ))}
            {filtered.length === 0 && (
              <TableRow>
                <TableCell colSpan={5} className="py-10 text-center text-muted-foreground">
                  No questions match your filters.
                </TableCell>
              </TableRow>
            )}
          </TableBody>
        </Table>
      </section>

      <ConfirmDialog
        open={!!pendingDelete}
        title="Delete this question?"
        message={`“${pendingDelete?.text.slice(0, 80) ?? ''}…” will be removed from the bank and all questionnaires that use it. This cannot be undone.`}
        confirmLabel="Delete question"
        onConfirm={confirmDelete}
        onCancel={() => setPendingDelete(null)}
      />
    </div>
  )
}
