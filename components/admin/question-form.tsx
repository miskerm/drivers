'use client'

import { useState } from 'react'
import { useRouter } from 'next/navigation'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { Select } from '@/components/ui/select'
import { Textarea } from '@/components/ui/textarea'
import { useModules } from '@/lib/module-store'
import type { Difficulty, ModuleId, QuestionType } from '@/lib/types'

export interface QuestionFormValues {
  moduleId: string
  type: QuestionType
  text: string
  options: string[]
  correct: number
  difficulty: Difficulty
  explanation: string
  image: string
}

export function QuestionForm({
  initial,
  submitLabel,
  onSubmit,
}: {
  initial?: QuestionFormValues
  submitLabel: string
  onSubmit: (values: QuestionFormValues) => void | Promise<void>
}) {
  const router = useRouter()
  const { modules } = useModules()
  const [values, setValues] = useState<QuestionFormValues>(
    initial ?? { moduleId: modules[0]?.id ?? 'traffic-signs', type: 'single', text: '', options: ['', '', '', ''], correct: 0, difficulty: 'easy', explanation: '', image: '' },
  )
  const [error, setError] = useState<string | null>(null)

  const set = <K extends keyof QuestionFormValues>(key: K, value: QuestionFormValues[K]) => setValues((v) => ({ ...v, [key]: value }))

  const setOption = (i: number, value: string) => {
    setValues((v) => ({ ...v, options: v.options.map((o, oi) => (oi === i ? value : o)) }))
  }

  const submit = async (e: React.FormEvent) => {
    e.preventDefault()
    const questionText = values.text.trim()
    const options = values.type === 'truefalse' ? ['True', 'False'] : values.options.map((option) => option.trim()).filter(Boolean)

    if (!questionText) return setError('Question text is required.')
    if (options.length < 2) return setError('Provide at least 2 answer options.')
    if (!options[values.correct]?.trim()) return setError('The marked correct answer must be a non-empty option.')

    setError(null)
    try {
      await onSubmit({ ...values, text: questionText, options })
      router.push('/admin/questions')
    } catch {
      setError('Could not save the question. Please check the database connection and try again.')
    }
  }

  return (
    <form onSubmit={submit} className="flex max-w-2xl flex-col gap-5">
      <div className="grid gap-4 sm:grid-cols-3">
        <div className="flex flex-col gap-2">
          <Label htmlFor="moduleId">Module</Label>
          <Select id="moduleId" value={values.moduleId} onChange={(e) => set('moduleId', e.target.value)}>
            {modules.map((m) => (
              <option key={m.id} value={m.id}>
                {m.title}
              </option>
            ))}
          </Select>
        </div>
        <div className="flex flex-col gap-2">
          <Label htmlFor="type">Question type</Label>
          <Select id="type" value={values.type} onChange={(e) => set('type', e.target.value as QuestionType)}>
            <option value="single">Single choice</option>
            <option value="truefalse">True / False</option>
          </Select>
        </div>
        <div className="flex flex-col gap-2">
          <Label htmlFor="difficulty">Difficulty</Label>
          <Select id="difficulty" value={values.difficulty} onChange={(e) => set('difficulty', e.target.value as Difficulty)}>
            <option value="easy">Easy</option>
            <option value="medium">Medium</option>
            <option value="hard">Hard</option>
          </Select>
        </div>
      </div>

      <div className="flex flex-col gap-2">
        <Label htmlFor="text">Question text</Label>
        <Textarea id="text" placeholder="What should you do at a stop sign?" value={values.text} onChange={(e) => set('text', e.target.value)} />
      </div>

      <div className="flex flex-col gap-3">
        <Label>Answer options — mark the correct one</Label>
        {(values.type === 'truefalse' ? ['True', 'False'] : values.options).map((option, i) => (
          <div key={i} className="flex items-center gap-2.5">
            <input
              type="radio"
              name="correct"
              checked={values.correct === i}
              onChange={() => set('correct', i)}
              aria-label={`Mark option ${String.fromCharCode(65 + i)} correct`}
              className="size-4 accent-[var(--primary)]"
            />
            {values.type === 'truefalse' ? (
              <p className="flex-1 rounded-lg bg-muted/60 px-3 py-2.5 text-sm font-semibold">{option}</p>
            ) : (
              <Input placeholder={`Option ${String.fromCharCode(65 + i)}`} value={option} onChange={(e) => setOption(i, e.target.value)} />
            )}
          </div>
        ))}
      </div>

      <div className="flex flex-col gap-2">
        <Label htmlFor="explanation">Explanation (shown after answering)</Label>
        <Textarea id="explanation" placeholder="Why is the correct answer right? What rule applies?" value={values.explanation} onChange={(e) => set('explanation', e.target.value)} />
      </div>

      <div className="flex flex-col gap-2">
        <Label htmlFor="image">Image URL (optional)</Label>
        <Input id="image" type="url" placeholder="https://…" value={values.image} onChange={(e) => set('image', e.target.value)} />
        <p className="text-xs text-muted-foreground">Uploaded files go to Cloudflare R2; store the object key here.</p>
      </div>

      {error && <p className="rounded-lg bg-destructive/10 px-3.5 py-2.5 text-sm font-medium text-destructive">{error}</p>}
      <div className="flex gap-2.5">
        <Button type="submit" className="h-10 px-5 text-sm">
          {submitLabel}
        </Button>
        <Button type="button" variant="outline" className="h-10 px-5 text-sm" onClick={() => router.push('/admin/questions')}>
          Cancel
        </Button>
      </div>
    </form>
  )
}
