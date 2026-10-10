'use client'

import { QuestionForm } from '@/components/admin/question-form'
import { createQuestion } from '@/lib/question-store'

export default function NewQuestionPage() {
  return (
    <div className="mx-auto max-w-4xl">
      <header className="mb-6">
        <h1 className="text-2xl font-extrabold tracking-tight sm:text-3xl">New question</h1>
        <p className="mt-1.5 text-sm text-muted-foreground">Add a question to the bank. It becomes available to questionnaires immediately.</p>
      </header>
      <div className="rounded-2xl border border-border/70 bg-card p-6 shadow-[0_1px_3px_rgb(16_24_40/0.05)]">
        <QuestionForm submitLabel="Create question" onSubmit={async (values) => { await createQuestion({ ...values, correct: [values.correct] }) }} />
      </div>
    </div>
  )
}
