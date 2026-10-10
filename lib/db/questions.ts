import { db, execute, executeOne, executeRun } from './index'
import type { Question } from '@/lib/types'

function rowToQuestion(row: any): Question {
  return {
    id: row.id,
    moduleId: row.module_id,
    type: row.type,
    text: row.text,
    options: JSON.parse(row.options),
    correct: JSON.parse(row.correct),
    difficulty: row.difficulty,
    explanation: row.explanation,
  }
}

export async function getQuestions(): Promise<Question[]> {
  const rows = await execute<any>('SELECT * FROM questions ORDER BY module_id, id')
  return rows.map(rowToQuestion)
}

export async function getQuestionsFor(moduleId: string): Promise<Question[]> {
  const rows = await execute<any>('SELECT * FROM questions WHERE module_id = ? ORDER BY id', [moduleId])
  return rows.map(rowToQuestion)
}

export async function getQuestion(id: string): Promise<Question | undefined> {
  const row = await executeOne<any>('SELECT * FROM questions WHERE id = ?', [id])
  return row ? rowToQuestion(row) : undefined
}

export async function createQuestion(input: Omit<Question, 'id'>): Promise<Question> {
  const id = `q-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 7)}`
  await executeRun(
    `INSERT INTO questions (id, module_id, type, text, options, correct, difficulty, explanation)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
    [id, input.moduleId, input.type, input.text, JSON.stringify(input.options), JSON.stringify(input.correct), input.difficulty, input.explanation]
  )
  return { ...input, id }
}

export async function createQuestions(inputs: Omit<Question, 'id'>[]): Promise<Question[]> {
  const questions: Question[] = []
  for (const input of inputs) questions.push(await createQuestion(input))
  return questions
}

export async function updateQuestion(id: string, patch: Partial<Question>): Promise<void> {
  const sets: string[] = []
  const args: any[] = []
  if (patch.moduleId !== undefined) { sets.push('module_id = ?'); args.push(patch.moduleId) }
  if (patch.type !== undefined) { sets.push('type = ?'); args.push(patch.type) }
  if (patch.text !== undefined) { sets.push('text = ?'); args.push(patch.text) }
  if (patch.options !== undefined) { sets.push('options = ?'); args.push(JSON.stringify(patch.options)) }
  if (patch.correct !== undefined) { sets.push('correct = ?'); args.push(JSON.stringify(patch.correct)) }
  if (patch.difficulty !== undefined) { sets.push('difficulty = ?'); args.push(patch.difficulty) }
  if (patch.explanation !== undefined) { sets.push('explanation = ?'); args.push(patch.explanation) }
  if (sets.length === 0) return
  args.push(id)
  await executeRun(`UPDATE questions SET ${sets.join(', ')} WHERE id = ?`, args)
}

export async function deleteQuestion(id: string): Promise<void> {
  await executeRun('DELETE FROM questions WHERE id = ?', [id])
}

export async function deleteQuestionsForModule(moduleId: string): Promise<void> {
  await executeRun('DELETE FROM questions WHERE module_id = ?', [moduleId])
}
