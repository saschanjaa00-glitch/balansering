import { useEffect, useMemo, useRef, useState } from 'react'
import './App.css'

type SourceTable = {
  name: string
  titleLine: string
  betweenTitleAndRows: string[]
  rowsMarkerLine: string
  betweenRowsAndHeader: string[]
  headerLine: string
  columnsRaw: string[]
  columnsSanitized: string[]
  rows: string[][]
  trailingBlankLines: string[]
}

type SourceSegment =
  | {
      type: 'text'
      lines: string[]
    }
  | {
      type: 'table'
      tableName: string
    }

type SourceDocument = {
  newline: string
  hasTrailingNewline: boolean
  segments: SourceSegment[]
  tables: Record<string, SourceTable>
}

type Assignment = {
  subjectCode: string
  subjectName: string
  groupCode: string
  block: string
}

type StudentRecord = {
  id: string
  fullName: string
  email: string
  classGroup: string
  assignments: Assignment[]
}

type SubjectRecord = {
  code: string
  name: string
  studentCount: number
  blocks: string[]
}

type GroupBreakdownRecord = {
  subjectCode: string
  subjectName: string
  groupCode: string
  block: string
  studentCount: number
}

type BlockBreakdownRecord = {
  block: string
  subjectCount: number
  groupCount: number
  studentCount: number
}

type ParsedData = {
  students: StudentRecord[]
  subjects: SubjectRecord[]
  groupBreakdowns: GroupBreakdownRecord[]
  blockBreakdowns: BlockBreakdownRecord[]
  blocks: string[]
  tableNames: string[]
  sourceDocument: SourceDocument
  initialAssignmentKeysByStudent: Record<string, string[]>
  sourceFileName: string
  originalAvailableBlocksBySubject: Record<string, string[]>
}

type BalanceChange = {
  studentId: string
  studentName: string
  subjectCode: string
  subjectName: string
  fromGroupCode: string
  fromBlock: string
  toGroupCode: string
  toBlock: string
}

type CollisionError = {
  studentId: string
  studentName: string
  classGroup: string
  subjects: Array<{ subjectCode: string; subjectName: string; block: string }>
}

type BalanceResultRun = {
  id: string
  createdAt: string
  message: string
  changes: BalanceChange[]
  collisionErrors?: CollisionError[]
}

type HistorySnapshot = {
  parsedData: ParsedData
  balanceResults: BalanceChange[] | null
  balanceHistory: BalanceResultRun[]
  balanceMessage: string
  debugGroups: Array<{ key: string; count: number; maxCap: number; status: string }>
}



const SUFFIX_BLOCKS: Record<string, string> = {
  A: 'Blokk 1',
  B: 'Blokk 2',
  C: 'Blokk 3',
  D: 'Blokk 4',
}

const BLOCK_SUFFIX_BY_NAME: Record<string, string> = {
  'Blokk 1': 'A',
  'Blokk 2': 'B',
  'Blokk 3': 'C',
  'Blokk 4': 'D',
}

const CUSTOM_SUBJECT_BLOCKS: Record<string, string> = {
  '2MAR5': 'MATTE',
  '2MAP3': 'MATTE',
  '2MAS5': 'MATTE',
  '3TY5': 'Blokk 1',
}

const EXCLUDED_GROUPS = new Set<string>([
  '2BID5A',
])

const EXCLUDED_SUBJECTS = new Set<string>([
  '2MP3',
])

const EXCLUDED_CLASS_GROUPS = new Set<string>([
  'IDA',
  'IDB',
  '2IDA',
  '2IDB',
  '3IDA',
  '3IDB',
])

const EXCLUDED_SUBJECT_TITLES_IN_OVERVIEW = new Set<string>([
  'Matematikk S1',
  'Matematikk R1',
  'Toppidrett 2',
  'Toppidrett 3',
])

const SUBJECT_MAX_CAPACITY: Record<string, number> = {
  'Kjemi 1': 24,
  'Kjemi 2': 24,
  'Biologi 1': 28,
  'Biologi 2': 28,
}

const STORAGE_KEYS = {
  parsedData: 'novaschem.parsedData.v1',
  uiState: 'novaschem.uiState.v1',
  balanceHistory: 'novaschem.balanceHistory.v1',
}

const MAX_HISTORY_STATES = 25

type PersistedUiState = {
  selectedStudentId: string
  studentQuery: string
  subjectQuery: string
  blockFilter: string
  viewMode: 'students' | 'subjects' | 'blokkoversikt' | 'bytteoversikt'
  onlyBlokkfag: boolean
  showIncompleteBlocks: boolean
  showOverloadedStudents: boolean
  showBlockCollisions: boolean
  showDuplicateSubjects: boolean
  perFaggruppeSortBy?: 'blokk' | 'tittel' | 'students' | 'change'
}

type BytteSubjectVisibility = Record<string, { vg2: boolean; vg3: boolean }>

function loadFromLocalStorage<T>(key: string, fallbackValue: T): T {
  try {
    const rawValue = window.localStorage.getItem(key)
    if (!rawValue) {
      return fallbackValue
    }
    return JSON.parse(rawValue) as T
  } catch {
    return fallbackValue
  }
}

function saveToLocalStorage<T>(key: string, value: T): void {
  try {
    window.localStorage.setItem(key, JSON.stringify(value))
  } catch {
    // Ignore storage errors (private mode, quota limits, etc.)
  }
}

function removeFromLocalStorage(key: string): void {
  try {
    window.localStorage.removeItem(key)
  } catch {
    // Ignore storage errors (private mode, quota limits, etc.)
  }
}

function summarizeBalanceChanges(rawChanges: BalanceChange[]): BalanceChange[] {
  const byStudentAndSubject = new Map<string, BalanceChange>()
  const keyOrder: string[] = []

  rawChanges.forEach((change) => {
    const key = `${change.studentId}|${change.subjectCode}`
    const existing = byStudentAndSubject.get(key)

    if (!existing) {
      byStudentAndSubject.set(key, { ...change })
      keyOrder.push(key)
      return
    }

    // Keep first origin and latest destination within one run.
    byStudentAndSubject.set(key, {
      ...existing,
      studentName: change.studentName || existing.studentName,
      subjectName: change.subjectName || existing.subjectName,
      toGroupCode: change.toGroupCode,
      toBlock: change.toBlock,
    })
  })

  return keyOrder
    .map((key) => byStudentAndSubject.get(key))
    .filter((change): change is BalanceChange => Boolean(change))
    .filter((change) => !(change.fromBlock === change.toBlock && change.fromGroupCode === change.toGroupCode))
}

function isInverseBalanceChange(previous: BalanceChange, next: BalanceChange): boolean {
  return (
    previous.studentId === next.studentId &&
    previous.subjectCode === next.subjectCode &&
    previous.fromGroupCode === next.toGroupCode &&
    previous.fromBlock === next.toBlock &&
    previous.toGroupCode === next.fromGroupCode &&
    previous.toBlock === next.fromBlock
  )
}

function formatBalanceChangeText(change: BalanceChange): string {
  const fromGroup = change.fromGroupCode.trim()
  const fromBlock = change.fromBlock.trim()
  const toGroup = change.toGroupCode.trim()
  const toBlock = change.toBlock.trim()

  if (!fromGroup && !fromBlock && (toGroup || toBlock)) {
    return `Lagt til: ${toBlock || '-'} (${toGroup || '-'})`
  }

  if (!toGroup && !toBlock && (fromGroup || fromBlock)) {
    return `Fjernet: ${fromBlock || '-'} (${fromGroup || '-'})`
  }

  return `${fromBlock || '-'} (${fromGroup || '-'}) -> ${toBlock || '-'} (${toGroup || '-'})`
}

function formatTimestamp(iso: string): string {
  const date = new Date(iso)
  if (Number.isNaN(date.getTime())) {
    return iso
  }
  return new Intl.DateTimeFormat('nb-NO', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
  }).format(date)
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/gu, '&amp;')
    .replace(/</gu, '&lt;')
    .replace(/>/gu, '&gt;')
    .replace(/"/gu, '&quot;')
    .replace(/'/gu, '&#39;')
}

const WINDOWS_1252_EXTENDED_MAP: Record<number, number> = {
  0x20ac: 0x80,
  0x201a: 0x82,
  0x0192: 0x83,
  0x201e: 0x84,
  0x2026: 0x85,
  0x2020: 0x86,
  0x2021: 0x87,
  0x02c6: 0x88,
  0x2030: 0x89,
  0x0160: 0x8a,
  0x2039: 0x8b,
  0x0152: 0x8c,
  0x017d: 0x8e,
  0x2018: 0x91,
  0x2019: 0x92,
  0x201c: 0x93,
  0x201d: 0x94,
  0x2022: 0x95,
  0x2013: 0x96,
  0x2014: 0x97,
  0x02dc: 0x98,
  0x2122: 0x99,
  0x0161: 0x9a,
  0x203a: 0x9b,
  0x0153: 0x9c,
  0x017e: 0x9e,
  0x0178: 0x9f,
}

function encodeWindows1252(text: string): ArrayBuffer {
  const bytes: number[] = []

  for (const char of text) {
    const codePoint = char.codePointAt(0)
    if (codePoint === undefined) {
      continue
    }

    const extended = WINDOWS_1252_EXTENDED_MAP[codePoint]
    if (extended !== undefined) {
      bytes.push(extended)
      continue
    }

    if (codePoint <= 0xff) {
      bytes.push(codePoint)
      continue
    }

    // Fallback for characters outside Windows-1252.
    bytes.push(0x3f)
  }

  return Uint8Array.from(bytes).buffer
}

function getMaxCapacityForSubject(subjectName: string): number {
  return SUBJECT_MAX_CAPACITY[subjectName] || 30
}

function countNumberedBlockAssignments(student: StudentRecord): number {
  return student.assignments.filter((assignment) => /^Blokk [1-4]$/u.test(assignment.block)).length
}

function hasMissingSubjects(student: StudentRecord): boolean {
  const blocksWithSubjects = new Set<string>()
  student.assignments.forEach((assignment) => {
    if (/^Blokk [1-4]$/u.test(assignment.block)) {
      blocksWithSubjects.add(assignment.block)
    }
  })
  return blocksWithSubjects.size < 3
}

function hasTooManySubjects(student: StudentRecord): boolean {
  return countNumberedBlockAssignments(student) >= 4
}

function hasBlockCollisions(student: StudentRecord): boolean {
  const blockUsageCount = new Map<string, number>()

  student.assignments.forEach((assignment) => {
    if (!/^Blokk [1-4]$/u.test(assignment.block)) {
      return
    }
    const currentCount = blockUsageCount.get(assignment.block) || 0
    blockUsageCount.set(assignment.block, currentCount + 1)
  })

  return Array.from(blockUsageCount.values()).some((count) => count > 1)
}

function hasDuplicateSubjects(student: StudentRecord): boolean {
  const seenSubjectCodes = new Set<string>()
  for (const assignment of student.assignments) {
    if (seenSubjectCodes.has(assignment.subjectCode)) {
      return true
    }
    seenSubjectCodes.add(assignment.subjectCode)
  }
  return false
}

function createDefaultGroupCode(subjectCode: string, block: string): string {
  const suffix = BLOCK_SUFFIX_BY_NAME[block]
  return suffix ? `${subjectCode}${suffix}` : subjectCode
}

/**
 * Fixes pre-existing block collisions in place.
 * A collision is when a student has two subjects scheduled in the same block.
 * For each such collision, we try to move one of the subjects to a different
 * block that is free for this student and has capacity in a group.
 * Operates on the live groupOccupancy map so it correctly reflects any moves
 * already made during the same balancing pass.
 */
function fixCollisionsInPlace(
  students: StudentRecord[],
  groupOccupancy: Map<string, string[]>,
  changes: BalanceChange[],
  studentsLookup: Map<string, StudentRecord>,
  maxCapacityOffset: number = 0
): CollisionError[] {
  const unresolvable: CollisionError[] = []
  // Build a live view: studentId -> block -> [{subjectCode, subjectName, groupCode}]
  const studentBlocks = new Map<string, Map<string, Array<{ subjectCode: string; subjectName: string; groupCode: string }>>>()

  groupOccupancy.forEach((occupants, occKey) => {
    const [sc, gc, blk] = occKey.split('|')
    if (!/^Blokk [1-4]$/u.test(blk)) return
    const subjectName = students.flatMap((s) => s.assignments).find((a) => a.subjectCode === sc)?.subjectName || ''
    if (!subjectName) return

    for (const sid of occupants) {
      if (!studentBlocks.has(sid)) studentBlocks.set(sid, new Map())
      const blockMap = studentBlocks.get(sid)!
      if (!blockMap.has(blk)) blockMap.set(blk, [])
      blockMap.get(blk)!.push({ subjectCode: sc, subjectName, groupCode: gc })
    }
  })

  studentBlocks.forEach((blockMap, studentId) => {
    const student = studentsLookup.get(studentId) ?? students.find((s) => s.id === studentId)
    if (!student) return

    blockMap.forEach((subjects, collidingBlock) => {
      if (subjects.length < 2) return

      // Helper: commit a collision-fix move and update all live state.
      const commitMove = (subjectCode: string, subjectName: string, groupCode: string, otherKey: string, otherGroupCode: string, otherBlock: string): void => {
        changes.push({
          studentId,
          studentName: student.fullName,
          subjectCode,
          subjectName,
          fromGroupCode: groupCode,
          fromBlock: collidingBlock,
          toGroupCode: otherGroupCode,
          toBlock: otherBlock,
        })

        const fromKey = `${subjectCode}|${groupCode}|${collidingBlock}`
        const idx = groupOccupancy.get(fromKey)?.indexOf(studentId) ?? -1
        if (idx !== -1) groupOccupancy.get(fromKey)?.splice(idx, 1)
        if (!groupOccupancy.has(otherKey)) groupOccupancy.set(otherKey, [])
        groupOccupancy.get(otherKey)!.push(studentId)

        const inColliding = blockMap.get(collidingBlock)!
        const subjectIdx = inColliding.findIndex((s) => s.subjectCode === subjectCode && s.groupCode === groupCode)
        if (subjectIdx !== -1) inColliding.splice(subjectIdx, 1)
        if (!blockMap.has(otherBlock)) blockMap.set(otherBlock, [])
        blockMap.get(otherBlock)!.push({ subjectCode, subjectName, groupCode: otherGroupCode })
      }

      // Try to move each subject out of the colliding block (try last first).
      for (const { subjectCode, subjectName, groupCode } of [...subjects].reverse()) {
        const maxCap = getMaxCapacityForSubject(subjectName) + maxCapacityOffset

        // Pass 1: prefer a group within capacity.
        let moved = false
        for (const [otherKey, occupants] of groupOccupancy.entries()) {
          const [otherSubject, otherGroupCode, otherBlock] = otherKey.split('|')
          if (otherSubject !== subjectCode) continue
          if (otherBlock === collidingBlock) continue
          if (!canStudentMoveToBlock(student, otherBlock)) continue
          if (occupants.length >= maxCap) continue
          if (blockMap.has(otherBlock)) continue
          if (occupants.includes(studentId)) continue

          commitMove(subjectCode, subjectName, groupCode, otherKey, otherGroupCode, otherBlock)
          moved = true
          break
        }

        // Pass 2: no capacity-respecting slot found — resolving the collision is more
        // important than holding the cap, so pick the least-full group in a free block.
        if (!moved) {
          const candidates: Array<{ key: string; groupCode: string; block: string; size: number }> = []
          for (const [otherKey, occupants] of groupOccupancy.entries()) {
            const [otherSubject, otherGroupCode, otherBlock] = otherKey.split('|')
            if (otherSubject !== subjectCode) continue
            if (otherBlock === collidingBlock) continue
            if (!canStudentMoveToBlock(student, otherBlock)) continue
            if (blockMap.has(otherBlock)) continue
            if (occupants.includes(studentId)) continue
            candidates.push({ key: otherKey, groupCode: otherGroupCode, block: otherBlock, size: occupants.length })
          }
          if (candidates.length > 0) {
            candidates.sort((a, b) => a.size - b.size)
            const best = candidates[0]
            commitMove(subjectCode, subjectName, groupCode, best.key, best.groupCode, best.block)
            moved = true
          }
        }

        // If this block's collision is already resolved, stop trying other subjects.
        const remaining = blockMap.get(collidingBlock)
        if (!remaining || remaining.length < 2) break
      }

      // After trying all subjects, if the block still has a collision it is unresolvable.
      const stillColliding = blockMap.get(collidingBlock)
      if (stillColliding && stillColliding.length >= 2) {
        const classGroup = (student.classGroup || '').trim()
        const existing = unresolvable.find((e) => e.studentId === studentId)
        const errorSubjects = stillColliding.map((s) => ({ subjectCode: s.subjectCode, subjectName: s.subjectName, block: collidingBlock }))
        if (existing) {
          for (const es of errorSubjects) {
            if (!existing.subjects.some((x) => x.subjectCode === es.subjectCode && x.block === es.block)) {
              existing.subjects.push(es)
            }
          }
        } else {
          unresolvable.push({ studentId, studentName: student.fullName, classGroup, subjects: errorSubjects })
        }
      }
    })
  })

  return unresolvable
}

function canStudentMoveToBlock(student: StudentRecord, targetBlock: string): boolean {
  const classGroup = (student.classGroup || '').trim().toUpperCase()

  // VG2 students cannot be moved to Blokk 4.
  if (/^2/u.test(classGroup) && targetBlock === 'Blokk 4') {
    return false
  }

  // VG3 students cannot be moved to Blokk 1.
  if (/^3/u.test(classGroup) && targetBlock === 'Blokk 1') {
    return false
  }

  return true
}

function balanceGroupsWithOffset(
  students: StudentRecord[],
  availableGroups: GroupBreakdownRecord[] = [],
  maxCapacityOffset: number = 0,
  debugCallback?: (groups: Array<{ key: string; count: number; maxCap: number; status: string }>) => void
): { changes: BalanceChange[]; overcrowdedCount: number; partnerLookaheadMoves: number } {
  const changes: BalanceChange[] = []
  const debugGroups: Array<{ key: string; count: number; maxCap: number; status: string }> = []
  const studentsById = new Map<string, StudentRecord>(students.map((student) => [student.id, student]))
  let partnerLookaheadMoves = 0

  // Guard rails to keep look-ahead search responsive on large datasets.
  const LOOKAHEAD_MAX_TARGET_GROUPS = 12
  const LOOKAHEAD_MAX_PARTNERS_PER_GROUP = 20
  const LOOKAHEAD_MAX_ATTEMPTS = 1600
  const LOOKAHEAD_MAX_DEPTH1_STUDENTS = 120
  const LOOKAHEAD_MAX_DEPTH2_STUDENTS = 40
  let lookaheadAttempts = 0

  // Build group occupancy map including empty but available groups.
  const groupOccupancy = buildGroupOccupancy(students, availableGroups)

  // Find overcrowded groups
  const overcrowded: Array<{ key: string; count: number; studentIds: string[] }> = []
  groupOccupancy.forEach((studentIds, key) => {
    const [subjectCode] = key.split('|')
    const subjectName = students
      .flatMap((s) => s.assignments)
      .find((a) => a.subjectCode === subjectCode)?.subjectName || 'UKJENT'

    const maxCap = getMaxCapacityForSubject(subjectName) + maxCapacityOffset
    const status = studentIds.length > maxCap ? `OVERFULL (${studentIds.length} > ${maxCap})` : `OK (${studentIds.length} ≤ ${maxCap})`
    
    debugGroups.push({ key, count: studentIds.length, maxCap, status })

    if (subjectName !== 'UKJENT' && studentIds.length > maxCap) {
      overcrowded.push({ key, count: studentIds.length, studentIds })
    }
  })

  if (debugCallback) {
    debugCallback(debugGroups)
  }

  // Helper to try a specific swap type for a student
  const trySwapForStudent = (
    studentId: string,
    subjectCode: string,
    groupCode: string,
    block: string,
    student: StudentRecord,
    subjectName: string,
    maxCap: number,
    swapType: 'simple' | 'double' | 'triple'
  ): boolean => {
    const blockAssignments = new Map<string, { subjectCode: string; subjectName: string; groupCode: string }[]>()
    student.assignments.forEach((a) => {
      if (!blockAssignments.has(a.block)) {
        blockAssignments.set(a.block, [])
      }
      blockAssignments.get(a.block)?.push(a)
    })

    const blocks = Array.from(blockAssignments.keys())
    if (blocks.length === 0) return false

    if (swapType === 'simple') {
      // Try to move to a block we don't already use
      const blockSet = new Set(blocks)
      for (const [otherKey, occupants] of groupOccupancy.entries()) {
        const [otherSubject, otherGroupCode, otherBlock] = otherKey.split('|')
        if (otherSubject === subjectCode && !blockSet.has(otherBlock) && occupants.length < maxCap && canStudentMoveToBlock(student, otherBlock)) {
          changes.push({
            studentId,
            studentName: student.fullName,
            subjectCode,
            subjectName,
            fromGroupCode: groupCode,
            fromBlock: block,
            toGroupCode: otherGroupCode,
            toBlock: otherBlock,
          })

          const key = `${subjectCode}|${groupCode}|${block}`
          const idx = groupOccupancy.get(key)?.indexOf(studentId) ?? -1
          if (idx !== -1) groupOccupancy.get(key)?.splice(idx, 1)
          if (!groupOccupancy.has(otherKey)) groupOccupancy.set(otherKey, [])
          groupOccupancy.get(otherKey)?.push(studentId)
          return true
        }
      }
    } else if (swapType === 'double' || swapType === 'triple') {
      // Try N-way swaps
      const swapLength = swapType === 'double' ? 2 : 3
      if (tryNWaySwap(studentId, subjectCode, block, student, blockAssignments, swapLength)) {
        return true
      }
    }

    return false
  }

  // Helper to try N-way swaps
  const tryNWaySwap = (
    studentId: string,
    sourceSubjectCode: string,
    sourceBlock: string,
    student: StudentRecord,
    blockAssignments: Map<string, { subjectCode: string; subjectName: string; groupCode: string }[]>,
    swapLength: number
  ): boolean => {
    const blocks = Array.from(blockAssignments.keys())
    const sourceBlockIdx = blocks.indexOf(sourceBlock)
    if (sourceBlockIdx === -1) return false

    // Generate all possible cycles of the given length starting from source block
    const generateCycles = (start: number, length: number, visited: Set<number> = new Set()): number[][] => {
      if (length === 0) return [[]]
      if (length === 1) return [[start]]

      const result: number[][] = []
      for (let i = 0; i < blocks.length; i++) {
        if (i === start || visited.has(i)) continue
        const newVisited = new Set(visited)
        newVisited.add(i)
        const subCycles = generateCycles(i, length - 1, newVisited)
        for (const cycle of subCycles) {
          result.push([start, ...cycle])
        }
      }
      return result
    }

    const cycles = generateCycles(sourceBlockIdx, swapLength)

    for (const cycle of cycles) {
      let isValid = true
      const cycleChanges: BalanceChange[] = []

      for (let i = 0; i < cycle.length; i++) {
        const fromBlockIdx = cycle[i]
        const toBlockIdx = cycle[(i + 1) % cycle.length]
        const fromBlock = blocks[fromBlockIdx]
        const toBlock = blocks[toBlockIdx]

        if (!canStudentMoveToBlock(student, toBlock)) {
          isValid = false
          break
        }

        const assignments = blockAssignments.get(fromBlock) || []
        if (assignments.length === 0) {
          isValid = false
          break
        }

        let targetSubject = assignments[0]
        if (i === 0) {
          targetSubject = assignments.find((a) => a.subjectCode === sourceSubjectCode) || assignments[0]
        }

        const targetSubjectName = targetSubject.subjectName
        const targetMaxCap = getMaxCapacityForSubject(targetSubjectName) + maxCapacityOffset
        const groupsInToBlock: Array<{ key: string; currentSize: number }> = []

        groupOccupancy.forEach((occupants, key) => {
          const [subCode, , blockName] = key.split('|')
          if (subCode === targetSubject.subjectCode && blockName === toBlock) {
            groupsInToBlock.push({ key, currentSize: occupants.length })
          }
        })

        if (groupsInToBlock.length === 0 || Math.min(...groupsInToBlock.map((g) => g.currentSize)) >= targetMaxCap) {
          isValid = false
          break
        }

        const targetGroup = groupsInToBlock.find((g) => g.currentSize < targetMaxCap)
        if (!targetGroup) {
          isValid = false
          break
        }

        cycleChanges.push({
          studentId,
          studentName: student.fullName,
          subjectCode: targetSubject.subjectCode,
          subjectName: targetSubjectName,
          fromGroupCode: targetSubject.groupCode,
          fromBlock: fromBlock,
          toGroupCode: targetGroup.key.split('|')[1],
          toBlock: toBlock,
        })
      }

      if (isValid && cycleChanges.length === swapLength) {
        for (const change of cycleChanges) {
          changes.push(change)

          const fromKey = `${change.subjectCode}|${change.fromGroupCode}|${change.fromBlock}`
          const toKey = `${change.subjectCode}|${change.toGroupCode}|${change.toBlock}`

          const idx = groupOccupancy.get(fromKey)?.indexOf(studentId) ?? -1
          if (idx !== -1) groupOccupancy.get(fromKey)?.splice(idx, 1)
          if (!groupOccupancy.has(toKey)) groupOccupancy.set(toKey, [])
          groupOccupancy.get(toKey)?.push(studentId)
        }
        return true
      }
    }

    return false
  }

  // Helper to try a partner-based switch with look-ahead.
  // If a student is stuck, pick a student in an alternative group,
  // try to move that partner via N-way swap, and then move the stuck student into the freed spot.
  const tryDirectSwitchWithLookahead = (
    studentId: string,
    subjectCode: string,
    groupCode: string,
    block: string,
    student: StudentRecord,
    subjectName: string,
    maxCap: number,
    lookaheadDepth: number,
    visitedStudentIds: Set<string> = new Set()
  ): boolean => {
    if (lookaheadDepth < 1) {
      return false
    }

    const visited = new Set(visitedStudentIds)
    visited.add(studentId)

    const sourceKey = `${subjectCode}|${groupCode}|${block}`

    // Candidate target groups for the same subject in other blocks.
    const targetGroups = Array.from(groupOccupancy.entries())
      .filter(([otherKey]) => {
        const [otherSubject, , otherBlock] = otherKey.split('|')
        return otherSubject === subjectCode && otherBlock !== block
      })
      .map(([otherKey, occupants]) => ({
        key: otherKey,
        count: occupants.length,
      }))
      .sort((a, b) => a.count - b.count)
      .slice(0, LOOKAHEAD_MAX_TARGET_GROUPS)

    for (const target of targetGroups) {
      const [, targetGroupCode, targetBlock] = target.key.split('|')
      const targetOccupants = [...(groupOccupancy.get(target.key) || [])]
        .sort((aId, bId) => {
          const aBlocks = new Set((studentsById.get(aId)?.assignments || []).map((a) => a.block)).size
          const bBlocks = new Set((studentsById.get(bId)?.assignments || []).map((a) => a.block)).size
          // Fewer occupied blocks usually means higher chance to find an open destination block.
          return aBlocks - bBlocks
        })
        .slice(0, LOOKAHEAD_MAX_PARTNERS_PER_GROUP)

      for (const partnerId of targetOccupants) {
        if (lookaheadAttempts >= LOOKAHEAD_MAX_ATTEMPTS) {
          return false
        }
        lookaheadAttempts++

        if (partnerId === studentId) continue
        if (visited.has(partnerId)) continue

        const partner = studentsById.get(partnerId)
        if (!partner) continue

        // Keep partner look-ahead attempts transactional. If we cannot move
        // the original student into the freed slot, revert all tentative moves.
        const snapshotChangesLength = changes.length
        const snapshotOccupancy = new Map<string, string[]>()
        groupOccupancy.forEach((occupants, occupancyKey) => {
          snapshotOccupancy.set(occupancyKey, [...occupants])
        })

        const restoreSnapshot = (): void => {
          changes.splice(snapshotChangesLength)
          groupOccupancy.clear()
          snapshotOccupancy.forEach((occupants, occupancyKey) => {
            groupOccupancy.set(occupancyKey, [...occupants])
          })
        }

        // Look-ahead: give the partner a chance to escape via N-way swap first.
        let movedPartner =
          trySwapForStudent(partnerId, subjectCode, targetGroupCode, targetBlock, partner, subjectName, maxCap, 'double') ||
          trySwapForStudent(partnerId, subjectCode, targetGroupCode, targetBlock, partner, subjectName, maxCap, 'triple') ||
          trySwapForStudent(partnerId, subjectCode, targetGroupCode, targetBlock, partner, subjectName, maxCap, 'simple')

        if (!movedPartner && lookaheadDepth > 1) {
          // Recursive fallback: try to free the partner via another partner chain.
          movedPartner = tryDirectSwitchWithLookahead(
            partnerId,
            subjectCode,
            targetGroupCode,
            targetBlock,
            partner,
            subjectName,
            maxCap,
            lookaheadDepth - 1,
            visited
          )
        }

        if (!movedPartner) {
          restoreSnapshot()
          continue
        }

        const targetCountAfterPartnerMove = (groupOccupancy.get(target.key) || []).length
        if (targetCountAfterPartnerMove >= maxCap) {
          restoreSnapshot()
          continue
        }

        if (!canStudentMoveToBlock(student, targetBlock)) {
          restoreSnapshot()
          continue
        }

        // Block collision guard: automatic balancing must never place a student into a block
        // they already occupy with a different subject. (Manual overrides may still do this.)
        const studentAlreadyInTargetBlock = Array.from(groupOccupancy.entries()).some(
          ([occupancyKey, occupants]) => {
            const [otherSubjectCode, , blockName] = occupancyKey.split('|')
            return blockName === targetBlock && otherSubjectCode !== subjectCode && occupants.includes(studentId)
          }
        )
        if (studentAlreadyInTargetBlock) {
          restoreSnapshot()
          continue
        }

        // Move the stuck student into the newly freed target slot.
        changes.push({
          studentId,
          studentName: student.fullName,
          subjectCode,
          subjectName,
          fromGroupCode: groupCode,
          fromBlock: block,
          toGroupCode: targetGroupCode,
          toBlock: targetBlock,
        })

        const sourceIndex = groupOccupancy.get(sourceKey)?.indexOf(studentId) ?? -1
        if (sourceIndex !== -1) {
          groupOccupancy.get(sourceKey)?.splice(sourceIndex, 1)
        }
        if (!groupOccupancy.has(target.key)) {
          groupOccupancy.set(target.key, [])
        }
        groupOccupancy.get(target.key)?.push(studentId)

        return true
      }
    }

    return false
  }

  // Try to rebalance in phases: simple, then double, then triple,
  // then partner look-ahead switches (depth 1, fallback depth 2 if depth 1 fails).
  overcrowded.forEach(({ key, count, studentIds }) => {
    const [subjectCode, groupCode, block] = key.split('|')
    const subjectName = students.flatMap((s) => s.assignments).find((a) => a.subjectCode === subjectCode)?.subjectName || ''
    const maxCap = getMaxCapacityForSubject(subjectName) + maxCapacityOffset
    const needToMove = count - maxCap

    let moved = 0
    let remainingStudents = [...studentIds]

    // Phase 1: Try simple swaps for all students
    remainingStudents = remainingStudents.filter((studentId) => {
      if (moved >= needToMove) return true

      const student = studentsById.get(studentId)
      if (!student) return true

      if (trySwapForStudent(studentId, subjectCode, groupCode, block, student, subjectName, maxCap, 'simple')) {
        moved++
        return false
      }
      return true
    })

    // Phase 2: Try double swaps for remaining students
    remainingStudents = remainingStudents.filter((studentId) => {
      if (moved >= needToMove) return true

      const student = studentsById.get(studentId)
      if (!student) return true

      if (trySwapForStudent(studentId, subjectCode, groupCode, block, student, subjectName, maxCap, 'double')) {
        moved++
        return false
      }
      return true
    })

    // Phase 3: Try triple swaps for remaining students
    remainingStudents = remainingStudents.filter((studentId) => {
      if (moved >= needToMove) return true

      const student = studentsById.get(studentId)
      if (!student) return true

      if (trySwapForStudent(studentId, subjectCode, groupCode, block, student, subjectName, maxCap, 'triple')) {
        moved++
        return false
      }
      return true
    })

    // Phase 4a: Try direct partner switches with depth-1 look-ahead.
    let phase4Depth1Moves = 0
    remainingStudents = remainingStudents.filter((studentId, index) => {
      if (moved >= needToMove) return true
      if (index >= LOOKAHEAD_MAX_DEPTH1_STUDENTS) return true

      const student = studentsById.get(studentId)
      if (!student) return true

      if (tryDirectSwitchWithLookahead(studentId, subjectCode, groupCode, block, student, subjectName, maxCap, 1)) {
        moved++
        phase4Depth1Moves++
        partnerLookaheadMoves++
        return false
      }
      return true
    })

    // Phase 4b: If depth 1 made no progress at all, retry with depth 2.
    if (moved < needToMove && phase4Depth1Moves === 0) {
      remainingStudents.slice(0, LOOKAHEAD_MAX_DEPTH2_STUDENTS).forEach((studentId) => {
        if (moved >= needToMove) return

        const student = studentsById.get(studentId)
        if (!student) return

        if (tryDirectSwitchWithLookahead(studentId, subjectCode, groupCode, block, student, subjectName, maxCap, 2)) {
          moved++
          partnerLookaheadMoves++
        }
      })
    }
  })

  // Phase 5: Fix pre-existing block collisions (two subjects in the same block).
  const collisionErrors = fixCollisionsInPlace(students, groupOccupancy, changes, studentsById, maxCapacityOffset)

  return { changes, overcrowdedCount: overcrowded.length, partnerLookaheadMoves, collisionErrors }
}

function progressiveBalanceGroups(
  students: StudentRecord[],
  availableGroups: GroupBreakdownRecord[] = [],
  maxOffset: number,
  debugCallback?: (groups: Array<{ key: string; count: number; maxCap: number; status: string }>) => void
): { allChanges: BalanceChange[]; summary: string; collisionErrors: CollisionError[] } {
  const allChanges: BalanceChange[] = []
  const offsets = []
  let totalPartnerLookaheadMoves = 0
  const allCollisionErrors: CollisionError[] = []
  
  // Generate offsets from maxOffset down to 0
  for (let i = maxOffset; i <= 0; i++) {
    offsets.push(i)
  }

  // Apply balancing iteratively
  let currentStudents = students
  
  for (const offset of offsets) {
    const result = balanceGroupsWithOffset(currentStudents, availableGroups, offset, debugCallback)
    totalPartnerLookaheadMoves += result.partnerLookaheadMoves

    // Merge collision errors, keeping only those still unresolved.
    for (const err of result.collisionErrors) {
      if (!allCollisionErrors.some((e) => e.studentId === err.studentId)) {
        allCollisionErrors.push(err)
      }
    }
    
    if (result.changes.length > 0) {
      // Apply the changes to get the new student state for next iteration
      currentStudents = applyBalanceChanges(currentStudents, result.changes)
      allChanges.push(...result.changes)
    } else if (result.overcrowdedCount === 0) {
      // No more overcrowded groups at this offset, we can stop
      break
    }
  }

  const uniqueStudents = new Set(allChanges.map((c) => c.studentId))
  const summary = `Progressiv balansering fullført: ${uniqueStudents.size} elev(er) flyttet (${allChanges.length} fagendringer, ${totalPartnerLookaheadMoves} partner-lookahead)`

  return { allChanges, summary, collisionErrors: allCollisionErrors }
}

// Advanced timeslot-based balancing algorithm

function balanceGroups(students: StudentRecord[], debugCallback?: (groups: Array<{ key: string; count: number; maxCap: number; status: string }>) => void): { changes: BalanceChange[]; overcrowdedCount: number } {
  const changes: BalanceChange[] = []
  const debugGroups: Array<{ key: string; count: number; maxCap: number; status: string }> = []

  // Build group occupancy map: key = "subjectCode|groupCode|block", value = [studentIds]
  const groupOccupancy = new Map<string, string[]>()

  students.forEach((student) => {
    student.assignments.forEach((assignment) => {
      const key = `${assignment.subjectCode}|${assignment.groupCode}|${assignment.block}`
      if (!groupOccupancy.has(key)) {
        groupOccupancy.set(key, [])
      }
      groupOccupancy.get(key)?.push(student.id)
    })
  })

  // Find overcrowded groups
  const overcrowded: Array<{ key: string; count: number; studentIds: string[] }> = []
  groupOccupancy.forEach((studentIds, key) => {
    const [subjectCode] = key.split('|')
    const subjectName = students
      .flatMap((s) => s.assignments)
      .find((a) => a.subjectCode === subjectCode)?.subjectName || 'UKJENT'

    const maxCap = getMaxCapacityForSubject(subjectName)
    const status = studentIds.length > maxCap ? `OVERFULL (${studentIds.length} > ${maxCap})` : `OK (${studentIds.length} ≤ ${maxCap})`
    
    debugGroups.push({ key, count: studentIds.length, maxCap, status })

    if (subjectName !== 'UKJENT' && studentIds.length > maxCap) {
      overcrowded.push({ key, count: studentIds.length, studentIds })
    }
  })

  if (debugCallback) {
    debugCallback(debugGroups)
  }

  // Helper to try a specific swap type for a student
  const trySwapForStudent = (
    studentId: string,
    subjectCode: string,
    groupCode: string,
    block: string,
    student: StudentRecord,
    subjectName: string,
    maxCap: number,
    swapType: 'simple' | 'double' | 'triple'
  ): boolean => {
    const blockAssignments = new Map<string, { subjectCode: string; subjectName: string; groupCode: string }[]>()
    student.assignments.forEach((a) => {
      if (!blockAssignments.has(a.block)) {
        blockAssignments.set(a.block, [])
      }
      blockAssignments.get(a.block)?.push(a)
    })

    const blocks = Array.from(blockAssignments.keys())
    if (blocks.length === 0) return false

    if (swapType === 'simple') {
      // Try to move to a block we don't already use
      const blockSet = new Set(blocks)
      for (const [otherKey, occupants] of groupOccupancy.entries()) {
        const [otherSubject, otherGroupCode, otherBlock] = otherKey.split('|')
        if (otherSubject === subjectCode && !blockSet.has(otherBlock) && occupants.length < maxCap && canStudentMoveToBlock(student, otherBlock)) {
          changes.push({
            studentId,
            studentName: student.fullName,
            subjectCode,
            subjectName,
            fromGroupCode: groupCode,
            fromBlock: block,
            toGroupCode: otherGroupCode,
            toBlock: otherBlock,
          })

          const key = `${subjectCode}|${groupCode}|${block}`
          const idx = groupOccupancy.get(key)?.indexOf(studentId) ?? -1
          if (idx !== -1) groupOccupancy.get(key)?.splice(idx, 1)
          if (!groupOccupancy.has(otherKey)) groupOccupancy.set(otherKey, [])
          groupOccupancy.get(otherKey)?.push(studentId)
          return true
        }
      }
    } else if (swapType === 'double' || swapType === 'triple') {
      // Try N-way swaps
      const swapLength = swapType === 'double' ? 2 : 3
      if (tryNWaySwap(studentId, subjectCode, block, student, blockAssignments, swapLength)) {
        return true
      }
    }

    return false
  }

  // Helper to try N-way swaps
  const tryNWaySwap = (
    studentId: string,
    sourceSubjectCode: string,
    sourceBlock: string,
    student: StudentRecord,
    blockAssignments: Map<string, { subjectCode: string; subjectName: string; groupCode: string }[]>,
    swapLength: number
  ): boolean => {
    const blocks = Array.from(blockAssignments.keys())
    const sourceBlockIdx = blocks.indexOf(sourceBlock)
    if (sourceBlockIdx === -1) return false

    // Generate all possible cycles of the given length starting from source block
    const generateCycles = (start: number, length: number, visited: Set<number> = new Set()): number[][] => {
      if (length === 0) return [[]]
      if (length === 1) return [[start]]

      const result: number[][] = []
      for (let i = 0; i < blocks.length; i++) {
        if (i === start || visited.has(i)) continue
        const newVisited = new Set(visited)
        newVisited.add(i)
        const subCycles = generateCycles(i, length - 1, newVisited)
        for (const cycle of subCycles) {
          result.push([start, ...cycle])
        }
      }
      return result
    }

    const cycles = generateCycles(sourceBlockIdx, swapLength)

    for (const cycle of cycles) {
      let isValid = true
      const cycleChanges: BalanceChange[] = []

      for (let i = 0; i < cycle.length; i++) {
        const fromBlockIdx = cycle[i]
        const toBlockIdx = cycle[(i + 1) % cycle.length]
        const fromBlock = blocks[fromBlockIdx]
        const toBlock = blocks[toBlockIdx]

        if (!canStudentMoveToBlock(student, toBlock)) {
          isValid = false
          break
        }

        const assignments = blockAssignments.get(fromBlock) || []
        if (assignments.length === 0) {
          isValid = false
          break
        }

        let targetSubject = assignments[0]
        if (i === 0) {
          targetSubject = assignments.find((a) => a.subjectCode === sourceSubjectCode) || assignments[0]
        }

        const targetSubjectName = targetSubject.subjectName
        const targetMaxCap = getMaxCapacityForSubject(targetSubjectName)
        const groupsInToBlock: Array<{ key: string; currentSize: number }> = []

        groupOccupancy.forEach((occupants, key) => {
          const [subCode, , blockName] = key.split('|')
          if (subCode === targetSubject.subjectCode && blockName === toBlock) {
            groupsInToBlock.push({ key, currentSize: occupants.length })
          }
        })

        if (groupsInToBlock.length === 0 || Math.min(...groupsInToBlock.map((g) => g.currentSize)) >= targetMaxCap) {
          isValid = false
          break
        }

        const targetGroup = groupsInToBlock.find((g) => g.currentSize < targetMaxCap)
        if (!targetGroup) {
          isValid = false
          break
        }

        cycleChanges.push({
          studentId,
          studentName: student.fullName,
          subjectCode: targetSubject.subjectCode,
          subjectName: targetSubjectName,
          fromGroupCode: targetSubject.groupCode,
          fromBlock: fromBlock,
          toGroupCode: targetGroup.key.split('|')[1],
          toBlock: toBlock,
        })
      }

      if (isValid && cycleChanges.length === swapLength) {
        for (const change of cycleChanges) {
          changes.push(change)

          const fromKey = `${change.subjectCode}|${change.fromGroupCode}|${change.fromBlock}`
          const toKey = `${change.subjectCode}|${change.toGroupCode}|${change.toBlock}`

          const idx = groupOccupancy.get(fromKey)?.indexOf(studentId) ?? -1
          if (idx !== -1) groupOccupancy.get(fromKey)?.splice(idx, 1)
          if (!groupOccupancy.has(toKey)) groupOccupancy.set(toKey, [])
          groupOccupancy.get(toKey)?.push(studentId)
        }
        return true
      }
    }

    return false
  }

  // Try to rebalance in phases: simple, then double, then triple
  overcrowded.forEach(({ key, count, studentIds }) => {
    const [subjectCode, groupCode, block] = key.split('|')
    const subjectName = students.flatMap((s) => s.assignments).find((a) => a.subjectCode === subjectCode)?.subjectName || ''
    const maxCap = getMaxCapacityForSubject(subjectName)
    const needToMove = count - maxCap

    let moved = 0
    let remainingStudents = [...studentIds]

    // Phase 1: Try simple swaps for all students
    remainingStudents = remainingStudents.filter((studentId) => {
      if (moved >= needToMove) return true

      const student = students.find((s) => s.id === studentId)
      if (!student) return true

      if (trySwapForStudent(studentId, subjectCode, groupCode, block, student, subjectName, maxCap, 'simple')) {
        moved++
        return false
      }
      return true
    })

    // Phase 2: Try double swaps for remaining students
    remainingStudents = remainingStudents.filter((studentId) => {
      if (moved >= needToMove) return true

      const student = students.find((s) => s.id === studentId)
      if (!student) return true

      if (trySwapForStudent(studentId, subjectCode, groupCode, block, student, subjectName, maxCap, 'double')) {
        moved++
        return false
      }
      return true
    })

    // Phase 3: Try triple swaps for remaining students
    remainingStudents.forEach((studentId) => {
      if (moved >= needToMove) return

      const student = students.find((s) => s.id === studentId)
      if (!student) return

      if (trySwapForStudent(studentId, subjectCode, groupCode, block, student, subjectName, maxCap, 'triple')) {
        moved++
      }
    })
  })

  // Phase 4: Fix pre-existing block collisions (two subjects in the same block).
  const studentsById = new Map<string, StudentRecord>(students.map((s) => [s.id, s]))
  const collisionErrors = fixCollisionsInPlace(students, groupOccupancy, changes, studentsById)

  return { changes, overcrowdedCount: overcrowded.length, collisionErrors }
}

function applyBalanceChanges(students: StudentRecord[], changes: BalanceChange[]): StudentRecord[] {
  // Apply each student's changes in order so chained moves land on the final destination.
  return students.map((student) => {
    const studentChanges = changes.filter((c) => c.studentId === student.id)
    if (studentChanges.length === 0) {
      return student
    }

    const updatedAssignments = student.assignments.map((assignment) => ({ ...assignment }))

    studentChanges.forEach((change) => {
      if (!canStudentMoveToBlock(student, change.toBlock)) {
        return
      }

      const assignmentIndex = updatedAssignments.findIndex(
        (assignment) =>
          assignment.subjectCode === change.subjectCode &&
          assignment.groupCode === change.fromGroupCode &&
          assignment.block === change.fromBlock,
      )

      if (assignmentIndex === -1) {
        return
      }

      updatedAssignments[assignmentIndex] = {
        ...updatedAssignments[assignmentIndex],
        groupCode: change.toGroupCode,
        block: change.toBlock,
      }
    })

    return {
      ...student,
      assignments: updatedAssignments,
    }
  })
}

function buildGroupOccupancy(
  students: StudentRecord[],
  existingGroups: GroupBreakdownRecord[] = [],
): Map<string, string[]> {
  const groupOccupancy = new Map<string, string[]>()

  existingGroups.forEach((group) => {
    const key = `${group.subjectCode}|${group.groupCode}|${group.block}`
    if (!groupOccupancy.has(key)) {
      groupOccupancy.set(key, [])
    }
  })

  students.forEach((student) => {
    student.assignments.forEach((assignment) => {
      const key = `${assignment.subjectCode}|${assignment.groupCode}|${assignment.block}`
      if (!groupOccupancy.has(key)) {
        groupOccupancy.set(key, [])
      }
      groupOccupancy.get(key)?.push(student.id)
    })
  })

  return groupOccupancy
}

function recalculateBreakdowns(
  students: StudentRecord[],
  existingGroups: GroupBreakdownRecord[] = [],
): { groupBreakdowns: GroupBreakdownRecord[]; blockBreakdowns: BlockBreakdownRecord[] } {
  const groupSummary = new Map<string, { subjectCode: string; subjectName: string; groupCode: string; block: string; students: Set<string> }>()

  existingGroups.forEach((group) => {
    const key = `${group.subjectCode}|${group.groupCode}|${group.block}`
    if (!groupSummary.has(key)) {
      groupSummary.set(key, {
        subjectCode: group.subjectCode,
        subjectName: group.subjectName,
        groupCode: group.groupCode,
        block: group.block,
        students: new Set<string>(),
      })
    }
  })

  students.forEach((student) => {
    student.assignments.forEach((assignment) => {
      const key = `${assignment.subjectCode}|${assignment.groupCode}|${assignment.block}`
      if (!groupSummary.has(key)) {
        groupSummary.set(key, {
          subjectCode: assignment.subjectCode,
          subjectName: assignment.subjectName,
          groupCode: assignment.groupCode,
          block: assignment.block,
          students: new Set<string>(),
        })
      }
      groupSummary.get(key)?.students.add(student.id)
    })
  })

  const groupBreakdowns: GroupBreakdownRecord[] = Array.from(groupSummary.values())
    .map((group) => ({
      subjectCode: group.subjectCode,
      subjectName: group.subjectName,
      groupCode: group.groupCode,
      block: group.block,
      studentCount: group.students.size,
    }))
    .sort((a, b) => {
      const blockCmp = a.block.localeCompare(b.block)
      if (blockCmp !== 0) {
        return blockCmp
      }
      const subjectCmp = a.subjectCode.localeCompare(b.subjectCode)
      if (subjectCmp !== 0) {
        return subjectCmp
      }
      return a.groupCode.localeCompare(b.groupCode)
    })

  const blockSummary = new Map<string, { subjects: Set<string>; groups: Set<string>; students: Set<string> }>()
  students.forEach((student) => {
    student.assignments.forEach((assignment) => {
      if (!assignment.block) {
        return
      }
      if (!blockSummary.has(assignment.block)) {
        blockSummary.set(assignment.block, {
          subjects: new Set<string>(),
          groups: new Set<string>(),
          students: new Set<string>(),
        })
      }
      const summary = blockSummary.get(assignment.block)
      summary?.subjects.add(assignment.subjectCode)
      summary?.groups.add(`${assignment.subjectCode}|${assignment.groupCode}`)
      summary?.students.add(student.id)
    })
  })

  const blockBreakdowns: BlockBreakdownRecord[] = Array.from(blockSummary.entries())
    .map(([block, summary]) => ({
      block,
      subjectCount: summary.subjects.size,
      groupCount: summary.groups.size,
      studentCount: summary.students.size,
    }))
    .sort((a, b) => a.block.localeCompare(b.block))

  return { groupBreakdowns, blockBreakdowns }
}

function inferBlockFromSuffix(code: string): string {
  const trimmed = code.trim()
  if (trimmed.length < 6) {
    return ''
  }
  const suffix = trimmed.slice(-1).toUpperCase()
  return SUFFIX_BLOCKS[suffix] || ''
}

function getCustomBlock(subjectCode: string): string {
  if (CUSTOM_SUBJECT_BLOCKS[subjectCode]) {
    return CUSTOM_SUBJECT_BLOCKS[subjectCode]
  }
  const baseCode = subjectCode.replace(/[A-D]$/u, '')
  return CUSTOM_SUBJECT_BLOCKS[baseCode] || ''
}

function sortBlocks(blocks: string[]): string[] {
  return blocks.sort((a, b) => {
    const aIsNumeric = /^Blokk \d+$/u.test(a)
    const bIsNumeric = /^Blokk \d+$/u.test(b)

    if (aIsNumeric && bIsNumeric) {
      const aNum = parseInt(a.split(' ')[1], 10)
      const bNum = parseInt(b.split(' ')[1], 10)
      return aNum - bNum
    }

    if (aIsNumeric) {
      return -1
    }
    if (bIsNumeric) {
      return 1
    }

    return a.localeCompare(b)
  })
}

function formatSubjectDisplayName(subjectName: string): string {
  const normalized = subjectName.trim().toLowerCase()

  if (normalized === 'internasjonal engelsk, skriftlig') {
    return 'Engelsk 1'
  }

  if (normalized === 'samfunnsfaglig engelsk, skriftlig') {
    return 'Engelsk 2'
  }

  return subjectName
}

function getFinalSubjectsByBlock(student?: StudentRecord): Array<{ blockNumber: string; subjects: string }> {
  if (!student) {
    return []
  }

  const subjectsByBlock = new Map<string, string[]>()

  student.assignments.forEach((assignment) => {
    if (!/^Blokk [1-4]$/u.test(assignment.block)) {
      return
    }

    if (!subjectsByBlock.has(assignment.block)) {
      subjectsByBlock.set(assignment.block, [])
    }

    const subjectLabel = formatSubjectDisplayName(assignment.subjectName || assignment.subjectCode)
    const blockSubjects = subjectsByBlock.get(assignment.block)
    if (blockSubjects && !blockSubjects.includes(subjectLabel)) {
      blockSubjects.push(subjectLabel)
    }
  })

  return sortBlocks(Array.from(subjectsByBlock.keys()))
    .map((block) => {
      const blockNumber = block.replace('Blokk ', '')
      const subjectLabels = [...(subjectsByBlock.get(block) || [])].sort((a, b) => a.localeCompare(b, 'nb-NO'))
      return {
        blockNumber,
        subjects: subjectLabels.join(' / '),
      }
    })
}

function getSubjectBlockChoices(
  parsedData: ParsedData,
  subjectCode: string,
): Array<{ block: string; studentCount: number; groupCount: number }> {
  const subject = parsedData.subjects.find((item) => item.code === subjectCode)
  if (!subject) {
    return []
  }

  const byBlock = new Map<string, { studentCount: number; groupCount: number }>()
  subject.blocks.forEach((block) => {
    byBlock.set(block, { studentCount: 0, groupCount: 0 })
  })

  parsedData.groupBreakdowns.forEach((group) => {
    if (group.subjectCode !== subjectCode) {
      return
    }

    if (!byBlock.has(group.block)) {
      byBlock.set(group.block, { studentCount: 0, groupCount: 0 })
    }

    const current = byBlock.get(group.block)
    if (!current) {
      return
    }

    current.studentCount += group.studentCount
    current.groupCount += 1
  })

  return sortBlocks(Array.from(byBlock.keys())).map((block) => {
    const current = byBlock.get(block) || { studentCount: 0, groupCount: 0 }
    return {
      block,
      studentCount: current.studentCount,
      groupCount: current.groupCount,
    }
  })
}

function getBalanceRunLabel(message: string): string {
  const trimmed = message.trim()

  if (!trimmed) {
    return 'Balansering'
  }

  if (trimmed.startsWith('Progressiv balansering')) {
    return 'Progressiv balansering'
  }

  if (trimmed.startsWith('Masseoppdatering')) {
    return 'Masseoppdatering'
  }

  if (trimmed.startsWith('Fant ') || trimmed.startsWith('✓ Ingen overfulle grupper')) {
    return 'Balansering'
  }

  if (trimmed.includes('lagt til') || trimmed.includes('fjernet')) {
    return 'Manuell endring'
  }

  return trimmed
}

function groupBalanceRunsBySubject(
  runs: Array<{ runId: string; createdAt: string; message: string; changes: BalanceChange[] }>,
): Array<{
  subjectCode: string
  subjectName: string
  changes: Array<{ runId: string; createdAt: string; message: string; change: BalanceChange }>
}> {
  const grouped = new Map<string, {
    subjectCode: string
    subjectName: string
    changes: Array<{ runId: string; createdAt: string; message: string; change: BalanceChange }>
  }>()

  runs.forEach((run) => {
    run.changes.forEach((change) => {
      const key = change.subjectCode
      if (!grouped.has(key)) {
        grouped.set(key, {
          subjectCode: change.subjectCode,
          subjectName: formatSubjectDisplayName(change.subjectName || change.subjectCode),
          changes: [],
        })
      }

      grouped.get(key)?.changes.push({
        runId: run.runId,
        createdAt: run.createdAt,
        message: run.message,
        change,
      })
    })
  })

  return Array.from(grouped.values())
    .map((subjectGroup) => ({
      ...subjectGroup,
      changes: subjectGroup.changes.sort((a, b) => new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime()),
    }))
    .sort((a, b) => a.subjectName.localeCompare(b.subjectName, 'nb-NO'))
}

function sanitizeHeader(value: string): string {
  return value.replace(/\s+\(\d+\)$/u, '').trim()
}

function decodeBestEffort(fileBuffer: ArrayBuffer): string {
  const utf8Text = new TextDecoder('utf-8').decode(fileBuffer)
  const win1252Text = new TextDecoder('windows-1252').decode(fileBuffer)
  const utf8BadChars = (utf8Text.match(/�/gu) || []).length
  const win1252BadChars = (win1252Text.match(/�/gu) || []).length
  return win1252BadChars < utf8BadChars ? win1252Text : utf8Text
}

function parseSourceDocument(rawText: string): SourceDocument {
  const newline = rawText.includes('\r\n') ? '\r\n' : '\n'
  const normalizedLines = rawText.replace(/\r\n/gu, '\n').split('\n')
  const tableHeaderRegex = /^([A-Za-z_]+)\s+\(\d+\)$/u

  const segments: SourceSegment[] = []
  const tables: Record<string, SourceTable> = {}
  let bufferedTextLines: string[] = []

  let index = 0
  while (index < normalizedLines.length) {
    const currentLine = normalizedLines[index]
    const headerMatch = currentLine.trim().match(tableHeaderRegex)

    if (!headerMatch) {
      bufferedTextLines.push(currentLine)
      index += 1
      continue
    }

    let rowsMarkerIndex = index + 1
    while (rowsMarkerIndex < normalizedLines.length && normalizedLines[rowsMarkerIndex].trim() !== '[Rows]') {
      if (tableHeaderRegex.test(normalizedLines[rowsMarkerIndex].trim())) {
        break
      }
      rowsMarkerIndex += 1
    }

    if (rowsMarkerIndex >= normalizedLines.length || normalizedLines[rowsMarkerIndex].trim() !== '[Rows]') {
      bufferedTextLines.push(currentLine)
      index += 1
      continue
    }

    let headerLineIndex = rowsMarkerIndex + 1
    while (headerLineIndex < normalizedLines.length && normalizedLines[headerLineIndex].trim() === '') {
      headerLineIndex += 1
    }

    if (headerLineIndex >= normalizedLines.length) {
      bufferedTextLines.push(...normalizedLines.slice(index))
      break
    }

    if (bufferedTextLines.length > 0) {
      segments.push({ type: 'text', lines: bufferedTextLines })
      bufferedTextLines = []
    }

    const tableName = headerMatch[1]
    const rowLines: string[] = []
    let rowIndex = headerLineIndex + 1
    while (rowIndex < normalizedLines.length) {
      const trimmedRow = normalizedLines[rowIndex].trim()
      if (trimmedRow === '' || tableHeaderRegex.test(trimmedRow)) {
        break
      }
      rowLines.push(normalizedLines[rowIndex])
      rowIndex += 1
    }

    const trailingBlankLines: string[] = []
    while (rowIndex < normalizedLines.length && normalizedLines[rowIndex].trim() === '') {
      trailingBlankLines.push(normalizedLines[rowIndex])
      rowIndex += 1
    }

    const columnsRaw = normalizedLines[headerLineIndex].split('\t')
    const table: SourceTable = {
      name: tableName,
      titleLine: normalizedLines[index],
      betweenTitleAndRows: normalizedLines.slice(index + 1, rowsMarkerIndex),
      rowsMarkerLine: normalizedLines[rowsMarkerIndex],
      betweenRowsAndHeader: normalizedLines.slice(rowsMarkerIndex + 1, headerLineIndex),
      headerLine: normalizedLines[headerLineIndex],
      columnsRaw,
      columnsSanitized: columnsRaw.map((header) => sanitizeHeader(header)),
      rows: rowLines.map((line) => line.split('\t')),
      trailingBlankLines,
    }

    tables[tableName] = table
    segments.push({ type: 'table', tableName })
    index = rowIndex
  }

  if (bufferedTextLines.length > 0) {
    segments.push({ type: 'text', lines: bufferedTextLines })
  }

  return {
    newline,
    hasTrailingNewline: /\r?\n$/u.test(rawText),
    segments,
    tables,
  }
}

function rowsToObjects(table?: SourceTable): Array<Record<string, string>> {
  if (!table) {
    return []
  }

  return table.rows.map((row) => {
    const objectRow: Record<string, string> = {}
    table.columnsSanitized.forEach((column, i) => {
      if (!column) {
        return
      }
      objectRow[column] = (row[i] || '').trim()
    })
    return objectRow
  })
}

function parseNovaschemExport(rawText: string, sourceFileName: string): ParsedData {
  const sourceDocument = parseSourceDocument(rawText)
  const tableNames = Object.keys(sourceDocument.tables).sort((a, b) => a.localeCompare(b))

  const studentRows = rowsToObjects(sourceDocument.tables.Student)
  const subjectRows = rowsToObjects(sourceDocument.tables.Subject)
  const groupRows = rowsToObjects(sourceDocument.tables.Group)
  const groupStudentRows = rowsToObjects(sourceDocument.tables.Group_Student)
  const taRows = rowsToObjects(sourceDocument.tables.TA)

  const studentsById = new Map<string, { fullName: string; email: string }>()
  studentRows.forEach((row) => {
    const id = row.Student
    if (!id) {
      return
    }
    const name = [row.LastName, row.FirstName].filter((value) => value && value.trim().length > 0).join(' ').trim()
    studentsById.set(id, {
      fullName: name || `Student ${id}`,
      email: row.EMail || '',
    })
  })

  const subjectsByCode = new Map<string, string>()
  subjectRows.forEach((row) => {
    const code = row.Subject
    if (!code) {
      return
    }
    subjectsByCode.set(code, row.FullText || code)
  })

  const groupToStudents = new Map<string, Set<string>>()
  const classGroups = new Set<string>()
  const pushStudentToGroup = (groupCode: string, studentId: string): void => {
    if (!groupCode || !studentId) {
      return
    }
    if (!groupToStudents.has(groupCode)) {
      groupToStudents.set(groupCode, new Set<string>())
    }
    groupToStudents.get(groupCode)?.add(studentId)
  }

  groupRows.forEach((row) => {
    if (row.Class === '1') {
      classGroups.add(row.Group)
    }
  })

  groupStudentRows.forEach((row) => {
    pushStudentToGroup(row.Group, row.Student)
  })

  groupRows.forEach((row) => {
    const groupCode = row.Group
    const studentList = row.Student
    if (!groupCode || !studentList) {
      return
    }
    studentList
      .split(',')
      .map((value) => value.trim())
      .filter((value) => /^\d+$/u.test(value))
      .forEach((studentId) => pushStudentToGroup(groupCode, studentId))
  })

  const assignmentsByStudent = new Map<string, Map<string, Assignment>>()
  const taSubjectGroupPairs = new Set<string>()
  const availableSubjectBlocks = new Map<string, { name: string; blocks: Set<string> }>()
  const addAvailableSubjectBlock = (subjectCode: string, block: string): void => {
    if (!subjectCode || !block || EXCLUDED_SUBJECTS.has(subjectCode)) {
      return
    }
    if (!availableSubjectBlocks.has(subjectCode)) {
      availableSubjectBlocks.set(subjectCode, {
        name: subjectsByCode.get(subjectCode) || subjectCode,
        blocks: new Set<string>(),
      })
    }
    availableSubjectBlocks.get(subjectCode)?.blocks.add(block)
  }

  const addAssignment = (studentId: string, assignment: Assignment): void => {
    if (!assignmentsByStudent.has(studentId)) {
      assignmentsByStudent.set(studentId, new Map<string, Assignment>())
    }
    const key = `${assignment.subjectCode}|${assignment.groupCode}|${assignment.block}`
    assignmentsByStudent.get(studentId)?.set(key, assignment)
  }

  taRows.forEach((row) => {
    const subjectCode = row.Subject
    const groupCode = row.Group
    if (!subjectCode || !groupCode) {
      return
    }
    if (EXCLUDED_GROUPS.has(groupCode) || EXCLUDED_SUBJECTS.has(subjectCode)) {
      return
    }
    const linkedStudents = groupToStudents.get(groupCode)
    if (!linkedStudents || linkedStudents.size === 0) {
      return
    }

    const customBlock = getCustomBlock(subjectCode)
    const inferredBlock = customBlock || inferBlockFromSuffix(groupCode) || inferBlockFromSuffix(subjectCode)
    const blockName = row.Blockname || inferredBlock
    addAvailableSubjectBlock(subjectCode, blockName)

    const assignment: Assignment = {
      subjectCode,
      subjectName: subjectsByCode.get(subjectCode) || subjectCode,
      groupCode,
      block: blockName,
    }

    taSubjectGroupPairs.add(`${subjectCode}|${groupCode}`)

    linkedStudents.forEach((studentId) => {
      addAssignment(studentId, assignment)
    })
  })

  groupToStudents.forEach((studentIds, groupCode) => {
    if (EXCLUDED_GROUPS.has(groupCode)) {
      return
    }
    const suffixMatch = groupCode.match(/^(.+)([A-D])$/u)
    if (!suffixMatch) {
      return
    }

    const subjectCode = suffixMatch[1]
    if (EXCLUDED_SUBJECTS.has(subjectCode)) {
      return
    }
    if (!subjectsByCode.has(subjectCode)) {
      return
    }

    const pairKey = `${subjectCode}|${groupCode}`
    if (taSubjectGroupPairs.has(pairKey)) {
      return
    }

    const customBlock = getCustomBlock(subjectCode)
    const inferredBlock = customBlock || inferBlockFromSuffix(groupCode)
    addAvailableSubjectBlock(subjectCode, inferredBlock)

    const assignment: Assignment = {
      subjectCode,
      subjectName: subjectsByCode.get(subjectCode) || subjectCode,
      groupCode,
      block: inferredBlock,
    }

    studentIds.forEach((studentId) => {
      addAssignment(studentId, assignment)
    })
  })

  groupRows.forEach((row) => {
    const groupCode = row.Group
    if (!groupCode || EXCLUDED_GROUPS.has(groupCode)) {
      return
    }

    const suffixMatch = groupCode.match(/^(.+)([A-D])$/u)
    if (!suffixMatch) {
      return
    }

    const subjectCode = suffixMatch[1]
    if (EXCLUDED_SUBJECTS.has(subjectCode) || !subjectsByCode.has(subjectCode)) {
      return
    }

    const customBlock = getCustomBlock(subjectCode)
    const inferredBlock = customBlock || inferBlockFromSuffix(groupCode)
    addAvailableSubjectBlock(subjectCode, inferredBlock)
  })

  const studentClassMap = new Map<string, string>()
  groupRows.forEach((row) => {
    if (row.Class === '1') {
      const studentList = row.Student
      if (!studentList) {
        return
      }
      studentList
        .split(',')
        .map((value) => value.trim())
        .filter((value) => /^\d+$/u.test(value))
        .forEach((studentId) => {
          studentClassMap.set(studentId, row.Group)
        })
    }
  })

  const students: StudentRecord[] = Array.from(new Set([...studentsById.keys(), ...assignmentsByStudent.keys()]))
    .map((studentId) => {
      const studentMeta = studentsById.get(studentId)
      const assignmentList = Array.from(assignmentsByStudent.get(studentId)?.values() || [])
      assignmentList.sort((a, b) => {
        const subjectCmp = a.subjectCode.localeCompare(b.subjectCode)
        if (subjectCmp !== 0) {
          return subjectCmp
        }
        return a.groupCode.localeCompare(b.groupCode)
      })

      const classGroup = studentClassMap.get(studentId) || ''
      return {
        id: studentId,
        fullName: studentMeta?.fullName || `Student ${studentId}`,
        email: studentMeta?.email || '',
        classGroup: classGroup,
        assignments: assignmentList,
      }
    })
    .filter((student) => !EXCLUDED_CLASS_GROUPS.has(student.classGroup))
    .sort((a, b) => a.fullName.localeCompare(b.fullName))

  const subjectSummary = new Map<string, { code: string; name: string; students: Set<string>; blocks: Set<string> }>()
  const groupSummary = new Map<
    string,
    {
      subjectCode: string
      subjectName: string
      groupCode: string
      block: string
      students: Set<string>
    }
  >()

  students.forEach((student) => {
    student.assignments.forEach((assignment) => {
      if (!subjectSummary.has(assignment.subjectCode)) {
        subjectSummary.set(assignment.subjectCode, {
          code: assignment.subjectCode,
          name: assignment.subjectName,
          students: new Set<string>(),
          blocks: new Set<string>(),
        })
      }

      const summary = subjectSummary.get(assignment.subjectCode)
      summary?.students.add(student.id)
      if (assignment.block) {
        summary?.blocks.add(assignment.block)
      }

      const groupKey = `${assignment.subjectCode}|${assignment.groupCode}|${assignment.block}`
      if (!groupSummary.has(groupKey)) {
        groupSummary.set(groupKey, {
          subjectCode: assignment.subjectCode,
          subjectName: assignment.subjectName,
          groupCode: assignment.groupCode,
          block: assignment.block,
          students: new Set<string>(),
        })
      }
      groupSummary.get(groupKey)?.students.add(student.id)
    })
  })

  availableSubjectBlocks.forEach((item, subjectCode) => {
    if (!subjectSummary.has(subjectCode)) {
      subjectSummary.set(subjectCode, {
        code: subjectCode,
        name: item.name,
        students: new Set<string>(),
        blocks: new Set<string>(),
      })
    }

    const summary = subjectSummary.get(subjectCode)
    item.blocks.forEach((block) => {
      if (block) {
        summary?.blocks.add(block)
      }
    })
  })

  const subjects: SubjectRecord[] = Array.from(subjectSummary.values())
    .map((item) => ({
      code: item.code,
      name: item.name,
      studentCount: item.students.size,
      blocks: sortBlocks(Array.from(item.blocks)),
    }))
    .sort((a, b) => a.code.localeCompare(b.code))

  const blocks = sortBlocks(
    Array.from(
      new Set(
        subjects
          .flatMap((subject) => subject.blocks)
          .filter((blockName) => blockName.trim().length > 0),
      ),
    ),
  )

  const groupBreakdowns: GroupBreakdownRecord[] = Array.from(groupSummary.values())
    .map((group) => ({
      subjectCode: group.subjectCode,
      subjectName: group.subjectName,
      groupCode: group.groupCode,
      block: group.block,
      studentCount: group.students.size,
    }))
    .sort((a, b) => {
      const blockCmp = a.block.localeCompare(b.block)
      if (blockCmp !== 0) {
        return blockCmp
      }
      const subjectCmp = a.subjectCode.localeCompare(b.subjectCode)
      if (subjectCmp !== 0) {
        return subjectCmp
      }
      return a.groupCode.localeCompare(b.groupCode)
    })

  const blockSummary = new Map<string, { subjects: Set<string>; groups: Set<string>; students: Set<string> }>()
  students.forEach((student) => {
    student.assignments.forEach((assignment) => {
      if (!assignment.block) {
        return
      }
      if (!blockSummary.has(assignment.block)) {
        blockSummary.set(assignment.block, {
          subjects: new Set<string>(),
          groups: new Set<string>(),
          students: new Set<string>(),
        })
      }
      const summary = blockSummary.get(assignment.block)
      summary?.subjects.add(assignment.subjectCode)
      summary?.groups.add(`${assignment.subjectCode}|${assignment.groupCode}`)
      summary?.students.add(student.id)
    })
  })

  const blockBreakdowns: BlockBreakdownRecord[] = Array.from(blockSummary.entries())
    .map(([block, summary]) => ({
      block,
      subjectCount: summary.subjects.size,
      groupCount: summary.groups.size,
      studentCount: summary.students.size,
    }))
    .sort((a, b) => a.block.localeCompare(b.block))

  const initialAssignmentKeysByStudent: Record<string, string[]> = {}
  students.forEach((student) => {
    initialAssignmentKeysByStudent[student.id] = student.assignments
      .map((assignment) => `${assignment.subjectCode}|${assignment.groupCode}|${assignment.block}`)
      .sort((a, b) => a.localeCompare(b))
  })

  const originalAvailableBlocksBySubject: Record<string, string[]> = {}
  availableSubjectBlocks.forEach((item, subjectCode) => {
    originalAvailableBlocksBySubject[subjectCode] = sortBlocks(Array.from(item.blocks))
  })

  return {
    students,
    subjects,
    groupBreakdowns,
    blockBreakdowns,
    blocks,
    tableNames,
    sourceDocument,
    initialAssignmentKeysByStudent,
    sourceFileName,
    originalAvailableBlocksBySubject,
  }
}

function isParsedDataExportReady(data: ParsedData | null): data is ParsedData {
  if (!data) {
    return false
  }
  return Boolean(
    data.sourceDocument &&
      Array.isArray(data.sourceDocument.segments) &&
      data.sourceDocument.tables &&
      data.initialAssignmentKeysByStudent,
  )
}

function cloneTable(table: SourceTable): SourceTable {
  return {
    ...table,
    betweenTitleAndRows: [...table.betweenTitleAndRows],
    betweenRowsAndHeader: [...table.betweenRowsAndHeader],
    columnsRaw: [...table.columnsRaw],
    columnsSanitized: [...table.columnsSanitized],
    rows: table.rows.map((row) => [...row]),
    trailingBlankLines: [...table.trailingBlankLines],
  }
}

function parseAssignmentKey(key: string): { subjectCode: string; groupCode: string; block: string } {
  const [subjectCode = '', groupCode = '', block = ''] = key.split('|')
  return { subjectCode, groupCode, block }
}

function parseStudentCsv(value: string): string[] {
  return value
    .split(',')
    .map((item) => item.trim())
    .filter((item) => /^\d+$/u.test(item))
}

function computeTotalDeltaCounts(parsedData: ParsedData): {
  groupDeltas: Map<string, number>
  blockDeltas: Map<string, number>
} {
  const originalGroupCounts = new Map<string, number>()
  const originalBlockCounts = new Map<string, number>()

  Object.values(parsedData.initialAssignmentKeysByStudent).forEach((assignmentKeys) => {
    assignmentKeys.forEach((key) => {
      const { subjectCode, groupCode, block } = parseAssignmentKey(key)
      const groupKey = `${subjectCode}|${groupCode}|${block}`
      originalGroupCounts.set(groupKey, (originalGroupCounts.get(groupKey) || 0) + 1)

      if (block) {
        originalBlockCounts.set(block, (originalBlockCounts.get(block) || 0) + 1)
      }
    })
  })

  const currentGroupCounts = new Map<string, number>()
  parsedData.groupBreakdowns.forEach((item) => {
    const key = `${item.subjectCode}|${item.groupCode}|${item.block}`
    currentGroupCounts.set(key, item.studentCount)
  })

  const currentBlockCounts = new Map<string, number>()
  parsedData.blockBreakdowns.forEach((item) => {
    currentBlockCounts.set(item.block, item.studentCount)
  })

  const groupDeltas = new Map<string, number>()
  const allGroupKeys = new Set<string>([...Array.from(originalGroupCounts.keys()), ...Array.from(currentGroupCounts.keys())])
  allGroupKeys.forEach((key) => {
    const delta = (currentGroupCounts.get(key) || 0) - (originalGroupCounts.get(key) || 0)
    if (delta !== 0) {
      groupDeltas.set(key, delta)
    }
  })

  const blockDeltas = new Map<string, number>()
  const allBlockKeys = new Set<string>([...Array.from(originalBlockCounts.keys()), ...Array.from(currentBlockCounts.keys())])
  allBlockKeys.forEach((key) => {
    const delta = (currentBlockCounts.get(key) || 0) - (originalBlockCounts.get(key) || 0)
    if (delta !== 0) {
      blockDeltas.set(key, delta)
    }
  })

  return { groupDeltas, blockDeltas }
}

function buildExportText(parsedData: ParsedData): string {
  const sourceDocument = parsedData.sourceDocument
  const tables: Record<string, SourceTable> = {}

  Object.entries(sourceDocument.tables).forEach(([name, table]) => {
    tables[name] = cloneTable(table)
  })

  const currentAssignmentKeysByStudent = new Map<string, Set<string>>()
  parsedData.students.forEach((student) => {
    currentAssignmentKeysByStudent.set(
      student.id,
      new Set(student.assignments.map((assignment) => `${assignment.subjectCode}|${assignment.groupCode}|${assignment.block}`)),
    )
  })

  const removalsByPair = new Map<string, number>()
  const additionsByPair = new Map<string, number>()
  const affectedGroups = new Set<string>()

  const allStudentIds = new Set<string>([
    ...Object.keys(parsedData.initialAssignmentKeysByStudent),
    ...Array.from(currentAssignmentKeysByStudent.keys()),
  ])

  allStudentIds.forEach((studentId) => {
    const originalKeys = new Set(parsedData.initialAssignmentKeysByStudent[studentId] || [])
    const currentKeys = currentAssignmentKeysByStudent.get(studentId) || new Set<string>()

    originalKeys.forEach((key) => {
      if (currentKeys.has(key)) {
        return
      }
      const { groupCode } = parseAssignmentKey(key)
      const pairKey = `${groupCode}|${studentId}`
      removalsByPair.set(pairKey, (removalsByPair.get(pairKey) || 0) + 1)
      affectedGroups.add(groupCode)
    })

    currentKeys.forEach((key) => {
      if (originalKeys.has(key)) {
        return
      }
      const { groupCode } = parseAssignmentKey(key)
      const pairKey = `${groupCode}|${studentId}`
      additionsByPair.set(pairKey, (additionsByPair.get(pairKey) || 0) + 1)
      affectedGroups.add(groupCode)
    })
  })

  const groupStudentTable = tables.Group_Student
  if (groupStudentTable) {
    const groupColumnIndex = groupStudentTable.columnsSanitized.findIndex((name) => name === 'Group')
    const studentColumnIndex = groupStudentTable.columnsSanitized.findIndex((name) => name === 'Student')

    if (groupColumnIndex >= 0 && studentColumnIndex >= 0) {
      const nextRows: string[][] = []
      const removalsRemaining = new Map(removalsByPair)
      groupStudentTable.rows.forEach((row) => {
        const pairKey = `${(row[groupColumnIndex] || '').trim()}|${(row[studentColumnIndex] || '').trim()}`
        const removalsLeft = removalsRemaining.get(pairKey) || 0
        if (removalsLeft > 0) {
          removalsRemaining.set(pairKey, removalsLeft - 1)
          return
        }
        nextRows.push(row)
      })

      additionsByPair.forEach((count, pairKey) => {
        const [groupCode, studentId] = pairKey.split('|')
        for (let i = 0; i < count; i += 1) {
          const newRow = new Array(groupStudentTable.columnsRaw.length).fill('')
          newRow[groupColumnIndex] = groupCode
          newRow[studentColumnIndex] = studentId
          nextRows.push(newRow)
        }
      })

      groupStudentTable.rows = nextRows
    }
  }

  const groupTable = tables.Group
  if (groupTable && affectedGroups.size > 0) {
    const groupColumnIndex = groupTable.columnsSanitized.findIndex((name) => name === 'Group')
    const studentColumnIndex = groupTable.columnsSanitized.findIndex((name) => name === 'Student')

    if (groupColumnIndex >= 0 && studentColumnIndex >= 0) {
      groupTable.rows = groupTable.rows.map((row) => {
        const groupCode = (row[groupColumnIndex] || '').trim()
        if (!affectedGroups.has(groupCode)) {
          return row
        }

        const nextStudents = parseStudentCsv(row[studentColumnIndex] || '')
        const studentSet = new Set(nextStudents)

        removalsByPair.forEach((count, pairKey) => {
          if (count <= 0) {
            return
          }
          const [pairGroupCode, studentId] = pairKey.split('|')
          if (pairGroupCode === groupCode) {
            studentSet.delete(studentId)
          }
        })

        additionsByPair.forEach((count, pairKey) => {
          if (count <= 0) {
            return
          }
          const [pairGroupCode, studentId] = pairKey.split('|')
          if (pairGroupCode === groupCode) {
            studentSet.add(studentId)
          }
        })

        const orderedStudents = [...studentSet].sort((a, b) => Number(a) - Number(b))
        const nextRow = [...row]
        nextRow[studentColumnIndex] = orderedStudents.join(',')
        return nextRow
      })
    }
  }

  const chunks: string[] = []
  sourceDocument.segments.forEach((segment) => {
    if (segment.type === 'text') {
      chunks.push(segment.lines.join(sourceDocument.newline))
      return
    }

    const table = tables[segment.tableName] || sourceDocument.tables[segment.tableName]
    const tableLines = [
      table.titleLine,
      ...table.betweenTitleAndRows,
      table.rowsMarkerLine,
      ...table.betweenRowsAndHeader,
      table.headerLine,
      ...table.rows.map((row) => row.join('\t')),
      ...table.trailingBlankLines,
    ]
    chunks.push(tableLines.join(sourceDocument.newline))
  })

  const combined = chunks.join(sourceDocument.newline)
  return sourceDocument.hasTrailingNewline ? `${combined}${sourceDocument.newline}` : combined
}

function App() {
  const [persistedUiState] = useState<PersistedUiState>(() =>
    loadFromLocalStorage<PersistedUiState>(STORAGE_KEYS.uiState, {
      selectedStudentId: '',
      studentQuery: '',
      subjectQuery: '',
      blockFilter: '',
      viewMode: 'students',
      onlyBlokkfag: true,
      showIncompleteBlocks: false,
      showOverloadedStudents: false,
      showBlockCollisions: false,
      showDuplicateSubjects: false,
    })
  )

  const [parsedData, setParsedData] = useState<ParsedData | null>(() => {
    const stored = loadFromLocalStorage<ParsedData | null>(STORAGE_KEYS.parsedData, null)
    return isParsedDataExportReady(stored) ? stored : null
  })
  const [selectedStudentId, setSelectedStudentId] = useState<string>(persistedUiState.selectedStudentId)
  const [studentQuery, setStudentQuery] = useState<string>(persistedUiState.studentQuery)
  const [subjectQuery, setSubjectQuery] = useState<string>(persistedUiState.subjectQuery)
  const [blockFilter, setBlockFilter] = useState<string>(persistedUiState.blockFilter)
  const [viewMode, setViewMode] = useState<'students' | 'subjects' | 'blokkoversikt' | 'bytteoversikt'>(persistedUiState.viewMode)
  const [selectedGroupKey, setSelectedGroupKey] = useState<string>('')
  const [onlyBlokkfag, setOnlyBlokkfag] = useState<boolean>(persistedUiState.onlyBlokkfag)
  const [showIncompleteBlocks, setShowIncompleteBlocks] = useState<boolean>(persistedUiState.showIncompleteBlocks)
  const [showOverloadedStudents, setShowOverloadedStudents] = useState<boolean>(persistedUiState.showOverloadedStudents)
  const [showBlockCollisions, setShowBlockCollisions] = useState<boolean>(persistedUiState.showBlockCollisions)
  const [showDuplicateSubjects, setShowDuplicateSubjects] = useState<boolean>(persistedUiState.showDuplicateSubjects)
  const [perFaggruppeSortBy, setPerFaggruppeSortBy] = useState<'blokk' | 'tittel' | 'students' | 'change'>(
    persistedUiState.perFaggruppeSortBy ?? 'blokk',
  )
  const [balanceResults, setBalanceResults] = useState<BalanceChange[] | null>(null)
  const [balanceHistory, setBalanceHistory] = useState<BalanceResultRun[]>(() => {
    const stored = loadFromLocalStorage<BalanceResultRun[]>(STORAGE_KEYS.balanceHistory, [])
    return Array.isArray(stored)
      ? stored.filter((run) => run && Array.isArray(run.changes) && typeof run.createdAt === 'string')
      : []
  })
  const [balanceMessage, setBalanceMessage] = useState<string>('')
  const [showBalanceMessageHistory, setShowBalanceMessageHistory] = useState<boolean>(false)
  const [debugGroups, setDebugGroups] = useState<Array<{ key: string; count: number; maxCap: number; status: string }>>([])
  const [balanceDeltaCounts, setBalanceDeltaCounts] = useState<Map<string, number>>(new Map())
  const [balanceBlockDeltaCounts, setBalanceBlockDeltaCounts] = useState<Map<string, number>>(new Map())
  const [undoStack, setUndoStack] = useState<HistorySnapshot[]>([])
  const [redoStack, setRedoStack] = useState<HistorySnapshot[]>([])
  const [errorMessage, setErrorMessage] = useState<string>('')
  const [selectedStudentsForMassUpdate, setSelectedStudentsForMassUpdate] = useState<Set<string>>(new Set())
  const [showMassUpdateDialog, setShowMassUpdateDialog] = useState<boolean>(false)
  const [pendingMassRemoval, setPendingMassRemoval] = useState<boolean>(false)
  const [massUpdateTargetSubject, setMassUpdateTargetSubject] = useState<string>('')
  const [massUpdateTargetBlock, setMassUpdateTargetBlock] = useState<string>('')
  const [showAddSubjectDialog, setShowAddSubjectDialog] = useState<boolean>(false)
  const [addSubjectTargetCode, setAddSubjectTargetCode] = useState<string>('')
  const [addSubjectTargetBlock, setAddSubjectTargetBlock] = useState<string>('')
  const [showStudentAddSubjectDialog, setShowStudentAddSubjectDialog] = useState<boolean>(false)
  const [studentAddSubjectCode, setStudentAddSubjectCode] = useState<string>('')
  const [studentAddSubjectBlock, setStudentAddSubjectBlock] = useState<string>('')
  const [showStudentSwapDialog, setShowStudentSwapDialog] = useState<boolean>(false)
  const [studentSwapAssignmentKey, setStudentSwapAssignmentKey] = useState<string>('')
  const [studentSwapTargetSubject, setStudentSwapTargetSubject] = useState<string>('')
  const [studentSwapTargetBlock, setStudentSwapTargetBlock] = useState<string>('')
  const [pendingRemovalAssignment, setPendingRemovalAssignment] = useState<string>('')
  const [showProgressiveBalanceDialog, setShowProgressiveBalanceDialog] = useState<boolean>(false)
  const [progressiveBalanceMaxOffset, setProgressiveBalanceMaxOffset] = useState<number>(-4)
  const [bytteSubjectVisibility, setBytteSubjectVisibility] = useState<BytteSubjectVisibility>({})
  const [bytteOptionsExpanded, setBytteOptionsExpanded] = useState<boolean>(false)
  const fileInputRef = useRef<HTMLInputElement | null>(null)
  const isHistoryNavigationRef = useRef<boolean>(false)
  const lastHistorySnapshotRef = useRef<HistorySnapshot | null>(null)

  const createHistorySnapshot = (): HistorySnapshot | null => {
    if (!parsedData) {
      return null
    }

    return {
      parsedData,
      balanceResults: balanceResults ? [...balanceResults] : null,
      balanceHistory: [...balanceHistory],
      balanceMessage,
      debugGroups: [...debugGroups],
    }
  }

  const clearStoredData = (): void => {
    removeFromLocalStorage(STORAGE_KEYS.parsedData)
    removeFromLocalStorage(STORAGE_KEYS.uiState)
    removeFromLocalStorage(STORAGE_KEYS.balanceHistory)

    setParsedData(null)
    setSelectedStudentId('')
    setStudentQuery('')
    setSubjectQuery('')
    setBlockFilter('')
    setViewMode('students')
    setSelectedGroupKey('')
    setOnlyBlokkfag(true)
    setShowIncompleteBlocks(false)
    setShowOverloadedStudents(false)
    setShowBlockCollisions(false)
    setShowDuplicateSubjects(false)
    setPerFaggruppeSortBy('blokk')
    setBalanceResults(null)
    setBalanceHistory([])
    setBalanceMessage('')
    setDebugGroups([])
    setBalanceDeltaCounts(new Map())
    setBalanceBlockDeltaCounts(new Map())
    setUndoStack([])
    setRedoStack([])
    setErrorMessage('')
    setSelectedStudentsForMassUpdate(new Set())
    setShowMassUpdateDialog(false)
    setPendingMassRemoval(false)
    setMassUpdateTargetSubject('')
    setMassUpdateTargetBlock('')
    setShowAddSubjectDialog(false)
    setAddSubjectTargetCode('')
    setAddSubjectTargetBlock('')
    setShowStudentAddSubjectDialog(false)
    setStudentAddSubjectCode('')
    setStudentAddSubjectBlock('')
    setShowStudentSwapDialog(false)
    setStudentSwapAssignmentKey('')
    setStudentSwapTargetSubject('')
    setStudentSwapTargetBlock('')
    setPendingRemovalAssignment('')
    setBytteSubjectVisibility({})
    lastHistorySnapshotRef.current = null
    isHistoryNavigationRef.current = false

    // Allow selecting the same file again after clearing data.
    if (fileInputRef.current) {
      fileInputRef.current.value = ''
    }
  }

  useEffect(() => {
    saveToLocalStorage(STORAGE_KEYS.parsedData, parsedData)
  }, [parsedData])

  useEffect(() => {
    const currentSnapshot = createHistorySnapshot()

    if (isHistoryNavigationRef.current) {
      isHistoryNavigationRef.current = false
      lastHistorySnapshotRef.current = currentSnapshot
      return
    }

    if (!currentSnapshot) {
      setUndoStack([])
      setRedoStack([])
      lastHistorySnapshotRef.current = null
      return
    }

    const previousSnapshot = lastHistorySnapshotRef.current
    if (previousSnapshot && previousSnapshot.parsedData !== currentSnapshot.parsedData) {
      setUndoStack((stack) => [...stack, previousSnapshot].slice(-MAX_HISTORY_STATES))
      setRedoStack([])
    }

    lastHistorySnapshotRef.current = currentSnapshot
  }, [parsedData, balanceResults, balanceHistory, balanceMessage, debugGroups])

  const handleUndo = (): void => {
    if (undoStack.length === 0) {
      return
    }

    const previousState = undoStack[undoStack.length - 1]
    const currentSnapshot = createHistorySnapshot()
    setUndoStack((stack) => stack.slice(0, -1))
    if (currentSnapshot) {
      setRedoStack((stack) => [...stack, currentSnapshot].slice(-MAX_HISTORY_STATES))
    }
    isHistoryNavigationRef.current = true
    setParsedData(previousState.parsedData)
    setBalanceResults(previousState.balanceResults)
    setBalanceHistory(previousState.balanceHistory)
    setBalanceMessage(previousState.balanceMessage)
    setDebugGroups(previousState.debugGroups)
  }

  const handleRedo = (): void => {
    if (redoStack.length === 0) {
      return
    }

    const nextState = redoStack[redoStack.length - 1]
    const currentSnapshot = createHistorySnapshot()
    setRedoStack((stack) => stack.slice(0, -1))
    if (currentSnapshot) {
      setUndoStack((stack) => [...stack, currentSnapshot].slice(-MAX_HISTORY_STATES))
    }
    isHistoryNavigationRef.current = true
    setParsedData(nextState.parsedData)
    setBalanceResults(nextState.balanceResults)
    setBalanceHistory(nextState.balanceHistory)
    setBalanceMessage(nextState.balanceMessage)
    setDebugGroups(nextState.debugGroups)
  }

  useEffect(() => {
    if (!parsedData) {
      setBalanceDeltaCounts(new Map())
      setBalanceBlockDeltaCounts(new Map())
      return
    }

    const { groupDeltas, blockDeltas } = computeTotalDeltaCounts(parsedData)
    setBalanceDeltaCounts(groupDeltas)
    setBalanceBlockDeltaCounts(blockDeltas)
  }, [parsedData])

  useEffect(() => {
    if (!parsedData) {
      setBytteSubjectVisibility({})
      return
    }

    const subjectCodes = new Set(parsedData.subjects.map((subject) => subject.code))
    setBytteSubjectVisibility((previous) => {
      const next: BytteSubjectVisibility = {}

      for (const subject of parsedData.subjects) {
        const existing = previous[subject.code]
        next[subject.code] = existing ?? { vg2: true, vg3: true }
      }

      const hasSameEntries = Object.keys(previous).length === Object.keys(next).length
        && Array.from(subjectCodes).every((code) => {
          const prevValue = previous[code]
          const nextValue = next[code]
          return !!prevValue && prevValue.vg2 === nextValue.vg2 && prevValue.vg3 === nextValue.vg3
        })

      return hasSameEntries ? previous : next
    })
  }, [parsedData])

  useEffect(() => {
    saveToLocalStorage(STORAGE_KEYS.balanceHistory, balanceHistory)
  }, [balanceHistory])

  useEffect(() => {
    const stateToPersist: PersistedUiState = {
      selectedStudentId,
      studentQuery,
      subjectQuery,
      blockFilter,
      viewMode,
      onlyBlokkfag,
      showIncompleteBlocks,
      showOverloadedStudents,
      showBlockCollisions,
      showDuplicateSubjects,
      perFaggruppeSortBy,
    }
    saveToLocalStorage(STORAGE_KEYS.uiState, stateToPersist)
  }, [selectedStudentId, studentQuery, subjectQuery, blockFilter, viewMode, onlyBlokkfag, showIncompleteBlocks, showOverloadedStudents, showBlockCollisions, showDuplicateSubjects, perFaggruppeSortBy])

  const selectedStudent = useMemo(() => {
    if (!parsedData || !selectedStudentId) {
      return null
    }
    return parsedData.students.find((student) => student.id === selectedStudentId) || null
  }, [parsedData, selectedStudentId])

  const filteredStudents = useMemo(() => {
    if (!parsedData) {
      return []
    }
    const needle = studentQuery.trim().toLowerCase()
    return parsedData.students.filter((student) => {
      if (EXCLUDED_CLASS_GROUPS.has(student.classGroup)) {
        return false
      }
      const matchesSearch = !needle || student.fullName.toLowerCase().includes(needle) || student.id.includes(needle)
      if (!matchesSearch) {
        return false
      }

      if (showIncompleteBlocks) {
        if (!hasMissingSubjects(student)) {
          return false
        }
      }

      if (showOverloadedStudents && !hasTooManySubjects(student)) {
        return false
      }

      if (showBlockCollisions && !hasBlockCollisions(student)) {
        return false
      }

      if (showDuplicateSubjects && !hasDuplicateSubjects(student)) {
        return false
      }

      return true
    })
  }, [parsedData, studentQuery, showIncompleteBlocks, showOverloadedStudents, showBlockCollisions, showDuplicateSubjects])

  useEffect(() => {
    if (viewMode !== 'students') {
      return
    }

    if (filteredStudents.length === 0) {
      if (selectedStudentId) {
        setSelectedStudentId('')
      }
      return
    }

    const selectedStudentStillVisible = filteredStudents.some((student) => student.id === selectedStudentId)
    if (!selectedStudentStillVisible) {
      setSelectedStudentId(filteredStudents[0].id)
    }
  }, [viewMode, filteredStudents, selectedStudentId])

  const studentFilterCounts = useMemo(() => {
    if (!parsedData) {
      return { missingSubjectsCount: 0, tooManySubjectsCount: 0, blockCollisionsCount: 0, duplicateSubjectsCount: 0 }
    }

    const needle = studentQuery.trim().toLowerCase()
    const eligibleStudents = parsedData.students.filter((student) => {
      if (EXCLUDED_CLASS_GROUPS.has(student.classGroup)) {
        return false
      }
      return !needle || student.fullName.toLowerCase().includes(needle) || student.id.includes(needle)
    })

    return {
      missingSubjectsCount: eligibleStudents.filter((student) => hasMissingSubjects(student)).length,
      tooManySubjectsCount: eligibleStudents.filter((student) => hasTooManySubjects(student)).length,
      blockCollisionsCount: eligibleStudents.filter((student) => hasBlockCollisions(student)).length,
      duplicateSubjectsCount: eligibleStudents.filter((student) => hasDuplicateSubjects(student)).length,
    }
  }, [parsedData, studentQuery])

  const filteredAssignments = useMemo(() => {
    if (!selectedStudent) {
      return []
    }
    const needle = subjectQuery.trim().toLowerCase()
    const filtered = selectedStudent.assignments.filter((assignment) => {
      const matchesSubject =
        !needle ||
        assignment.subjectCode.toLowerCase().includes(needle) ||
        assignment.subjectName.toLowerCase().includes(needle)
      const matchesBlock = !blockFilter || assignment.block === blockFilter
      const isBlokkfag = assignment.block && assignment.block.trim().length > 0
      return matchesSubject && matchesBlock && (!onlyBlokkfag || isBlokkfag)
    })

    const getAssignmentBlockSortRank = (block: string): number => {
      const trimmed = block.trim()
      const numericMatch = trimmed.match(/^Blokk (\d+)$/u)
      if (numericMatch) {
        return parseInt(numericMatch[1], 10)
      }
      if (trimmed === 'MATTE') {
        return 999
      }
      return 500
    }

    return filtered.sort((a, b) => {
      const rankA = getAssignmentBlockSortRank(a.block)
      const rankB = getAssignmentBlockSortRank(b.block)
      if (rankA !== rankB) {
        return rankA - rankB
      }

      const blockCmp = a.block.localeCompare(b.block)
      if (blockCmp !== 0) {
        return blockCmp
      }

      return a.subjectName.localeCompare(b.subjectName)
    })
  }, [selectedStudent, subjectQuery, blockFilter, onlyBlokkfag])

  const filteredSubjects = useMemo(() => {
    if (!parsedData) {
      return []
    }
    const needle = subjectQuery.trim().toLowerCase()
    return parsedData.subjects
      .filter((subject) => {
        const matchesSearch =
          !needle || subject.code.toLowerCase().includes(needle) || subject.name.toLowerCase().includes(needle)
        const matchesBlock = !blockFilter || subject.blocks.includes(blockFilter)
        const isBlokkfag = subject.blocks && subject.blocks.length > 0
        return matchesSearch && matchesBlock && (!onlyBlokkfag || isBlokkfag)
      })
      .sort((a, b) => {
        // Get primary block for each subject (first block)
        const aBlock = a.blocks.length > 0 ? a.blocks[0] : ''
        const bBlock = b.blocks.length > 0 ? b.blocks[0] : ''

        // Compare blocks using same logic as sortBlocks
        const aIsNumeric = /^Blokk \d+$/u.test(aBlock)
        const bIsNumeric = /^Blokk \d+$/u.test(bBlock)

        let blockComparison = 0
        if (aIsNumeric && bIsNumeric) {
          const aNum = parseInt(aBlock.split(' ')[1], 10)
          const bNum = parseInt(bBlock.split(' ')[1], 10)
          blockComparison = aNum - bNum
        } else if (aIsNumeric) {
          blockComparison = -1
        } else if (bIsNumeric) {
          blockComparison = 1
        } else {
          blockComparison = aBlock.localeCompare(bBlock)
        }

        // If blocks are equal, sort by name
        if (blockComparison !== 0) {
          return blockComparison
        }
        return a.name.localeCompare(b.name)
      })
  }, [parsedData, subjectQuery, blockFilter, onlyBlokkfag])

  const blokkoversiktBlocksBySubject = useMemo(() => {
    const bySubject = new Map<string, Set<string>>()

    if (!parsedData) {
      return bySubject
    }

    Object.values(parsedData.initialAssignmentKeysByStudent).forEach((assignmentKeys) => {
      assignmentKeys.forEach((key) => {
        const { subjectCode, block } = parseAssignmentKey(key)
        if (!subjectCode || !block) {
          return
        }

        if (!bySubject.has(subjectCode)) {
          bySubject.set(subjectCode, new Set<string>())
        }
        bySubject.get(subjectCode)?.add(block)
      })
    })

    parsedData.groupBreakdowns.forEach((group) => {
      if (!group.subjectCode || !group.block) {
        return
      }

      if (!bySubject.has(group.subjectCode)) {
        bySubject.set(group.subjectCode, new Set<string>())
      }
      bySubject.get(group.subjectCode)?.add(group.block)
    })

    return bySubject
  }, [parsedData])

  const perSubjectBlockColumns = useMemo(() => {
    if (!parsedData) {
      return [] as string[]
    }

    const uniqueBlocks = Array.from(
      new Set(parsedData.blocks.filter((block) => block.trim().length > 0 && block !== 'MATTE')),
    )
    return sortBlocks(uniqueBlocks)
  }, [parsedData])

  const perSubjectMatrixRows = useMemo(() => {
    if (!parsedData) {
      return [] as Array<{
        subject: SubjectRecord
        maxCap: number
        countsByBlock: Map<string, number>
        hasGroupByBlock: Map<string, boolean>
        total: number
        isOverfilled: boolean
      }>
    }

    const countsBySubjectBlock = new Map<string, number>()
    const hasGroupBySubjectBlock = new Set<string>()
    const overfilledSubjectCodes = new Set<string>()

    parsedData.groupBreakdowns.forEach((group) => {
      const key = `${group.subjectCode}|${group.block}`
      countsBySubjectBlock.set(key, (countsBySubjectBlock.get(key) || 0) + group.studentCount)
      hasGroupBySubjectBlock.add(key)

      if (group.studentCount > getMaxCapacityForSubject(group.subjectName)) {
        overfilledSubjectCodes.add(group.subjectCode)
      }
    })

    return filteredSubjects
      .slice()
      .filter((subject) => !EXCLUDED_SUBJECT_TITLES_IN_OVERVIEW.has(subject.name))
      .sort((a, b) => a.name.localeCompare(b.name, 'nb-NO'))
      .map((subject) => {
        const maxCap = getMaxCapacityForSubject(subject.name)
        const countsByBlock = new Map<string, number>()
        const hasGroupByBlock = new Map<string, boolean>()

        perSubjectBlockColumns.forEach((block) => {
          const subjectBlockKey = `${subject.code}|${block}`
          const count = countsBySubjectBlock.get(subjectBlockKey) || 0
          countsByBlock.set(block, count)
          hasGroupByBlock.set(block, hasGroupBySubjectBlock.has(subjectBlockKey))
        })

        const total = perSubjectBlockColumns.reduce((sum, block) => {
          if (!hasGroupByBlock.get(block)) {
            return sum
          }
          return sum + (countsByBlock.get(block) || 0)
        }, 0)

        return {
          subject,
          maxCap,
          countsByBlock,
          hasGroupByBlock,
          total,
          isOverfilled: overfilledSubjectCodes.has(subject.code),
        }
      })
  }, [parsedData, filteredSubjects, perSubjectBlockColumns])

  const filteredGroupBreakdowns = useMemo(() => {
    if (!parsedData) {
      return []
    }

    const compareBlocks = (aBlock: string, bBlock: string): number => {
      const aIsNumeric = /^Blokk \d+$/u.test(aBlock)
      const bIsNumeric = /^Blokk \d+$/u.test(bBlock)

      if (aIsNumeric && bIsNumeric) {
        const aNum = parseInt(aBlock.split(' ')[1], 10)
        const bNum = parseInt(bBlock.split(' ')[1], 10)
        return aNum - bNum
      }
      if (aIsNumeric) {
        return -1
      }
      if (bIsNumeric) {
        return 1
      }
      return aBlock.localeCompare(bBlock)
    }

    const needle = subjectQuery.trim().toLowerCase()
    return parsedData.groupBreakdowns
      .filter((item) => {
        const matchesSearch =
          !needle ||
          item.subjectCode.toLowerCase().includes(needle) ||
          item.subjectName.toLowerCase().includes(needle) ||
          item.groupCode.toLowerCase().includes(needle)
        const isBlokkfag = item.block && item.block.trim().length > 0
         const isStandardBlock = /^Blokk \d+$/u.test(item.block)
         const matchesBlock = !blockFilter || item.block === blockFilter
         // Only show standard blocks (Blokk 1, Blokk 2, etc.) unless a specific block filter is selected
         const shouldIncludeBlock = blockFilter ? matchesBlock : isStandardBlock
         return matchesSearch && shouldIncludeBlock && (!onlyBlokkfag || isBlokkfag)
      })
      .sort((a, b) => {
        const blockComparison = compareBlocks(a.block, b.block)
        const titleComparison = a.subjectName.localeCompare(b.subjectName, 'nb-NO')

        if (perFaggruppeSortBy === 'blokk') {
          if (blockComparison !== 0) {
            return blockComparison
          }
          if (titleComparison !== 0) {
            return titleComparison
          }
          return a.groupCode.localeCompare(b.groupCode)
        }

        if (perFaggruppeSortBy === 'tittel') {
          if (titleComparison !== 0) {
            return titleComparison
          }
          if (blockComparison !== 0) {
            return blockComparison
          }
          return a.groupCode.localeCompare(b.groupCode)
        }

        if (perFaggruppeSortBy === 'students') {
          if (b.studentCount !== a.studentCount) {
            return b.studentCount - a.studentCount
          }
          if (titleComparison !== 0) {
            return titleComparison
          }
          if (blockComparison !== 0) {
            return blockComparison
          }
          return a.groupCode.localeCompare(b.groupCode)
        }

        const aDelta = balanceDeltaCounts.get(`${a.subjectCode}|${a.groupCode}|${a.block}`) || 0
        const bDelta = balanceDeltaCounts.get(`${b.subjectCode}|${b.groupCode}|${b.block}`) || 0
        if (bDelta !== aDelta) {
          return bDelta - aDelta
        }
        if (titleComparison !== 0) {
          return titleComparison
        }
        if (blockComparison !== 0) {
          return blockComparison
        }
        return a.groupCode.localeCompare(b.groupCode)
      })
  }, [parsedData, subjectQuery, blockFilter, onlyBlokkfag, perFaggruppeSortBy, balanceDeltaCounts])

  const studentMetaById = useMemo(() => {
    const map = new Map<string, { id: string; fullName: string; classGroup: string }>()
    if (!parsedData) {
      return map
    }

    parsedData.students.forEach((student) => {
      map.set(student.id, {
        id: student.id,
        fullName: student.fullName,
        classGroup: student.classGroup || '',
      })
    })

    return map
  }, [parsedData])

  const currentGroupMemberLookup = useMemo(() => {
    if (!parsedData) {
      return new Map<string, Array<{ id: string; fullName: string; classGroup: string }>>()
    }

    const lookup = new Map<string, Map<string, { id: string; fullName: string; classGroup: string }>>()
    parsedData.students.forEach((student) => {
      student.assignments.forEach((assignment) => {
        const key = `${assignment.subjectCode}-${assignment.groupCode}-${assignment.block}`
        if (!lookup.has(key)) {
          lookup.set(key, new Map<string, { id: string; fullName: string; classGroup: string }>())
        }
        lookup.get(key)?.set(student.id, {
          id: student.id,
          fullName: student.fullName,
          classGroup: student.classGroup || '',
        })
      })
    })

    const sortedLookup = new Map<string, Array<{ id: string; fullName: string; classGroup: string }>>()
    lookup.forEach((studentMap, key) => {
      const sortedStudents = Array.from(studentMap.values()).sort((a, b) => a.fullName.localeCompare(b.fullName))
      sortedLookup.set(key, sortedStudents)
    })
    return sortedLookup
  }, [parsedData])

  const originalGroupMemberLookup = useMemo(() => {
    if (!parsedData) {
      return new Map<string, Array<{ id: string; fullName: string; classGroup: string }>>()
    }

    const lookup = new Map<string, Map<string, { id: string; fullName: string; classGroup: string }>>()

    Object.entries(parsedData.initialAssignmentKeysByStudent).forEach(([studentId, assignmentKeys]) => {
      const meta = studentMetaById.get(studentId) || {
        id: studentId,
        fullName: `Student ${studentId}`,
        classGroup: '',
      }

      assignmentKeys.forEach((assignmentKey) => {
        const { subjectCode, groupCode, block } = parseAssignmentKey(assignmentKey)
        const key = `${subjectCode}-${groupCode}-${block}`
        if (!lookup.has(key)) {
          lookup.set(key, new Map<string, { id: string; fullName: string; classGroup: string }>())
        }
        lookup.get(key)?.set(studentId, meta)
      })
    })

    const sortedLookup = new Map<string, Array<{ id: string; fullName: string; classGroup: string }>>()
    lookup.forEach((studentMap, key) => {
      const sortedStudents = Array.from(studentMap.values()).sort((a, b) => a.fullName.localeCompare(b.fullName))
      sortedLookup.set(key, sortedStudents)
    })

    return sortedLookup
  }, [parsedData, studentMetaById])

  const selectedGroupMembers = useMemo(() => {
    if (!selectedGroupKey) {
      return [] as Array<{ id: string; fullName: string; classGroup: string; status: 'unchanged' | 'added' | 'removed' }>
    }

    const currentMembers = currentGroupMemberLookup.get(selectedGroupKey) || []
    const originalMembers = originalGroupMemberLookup.get(selectedGroupKey) || []
    const currentIds = new Set(currentMembers.map((student) => student.id))
    const originalIds = new Set(originalMembers.map((student) => student.id))

    const unchanged = currentMembers
      .filter((student) => originalIds.has(student.id))
      .map((student) => ({ ...student, status: 'unchanged' as const }))

    const added = currentMembers
      .filter((student) => !originalIds.has(student.id))
      .map((student) => ({ ...student, status: 'added' as const }))

    const removed = originalMembers
      .filter((student) => !currentIds.has(student.id))
      .map((student) => ({ ...student, status: 'removed' as const }))

    return [...unchanged, ...added, ...removed]
  }, [currentGroupMemberLookup, originalGroupMemberLookup, selectedGroupKey])

  const selectedGroupActiveMemberCount = useMemo(
    () => selectedGroupMembers.filter((student) => student.status !== 'removed').length,
    [selectedGroupMembers],
  )

  useEffect(() => {
    const activeIds = new Set(
      selectedGroupMembers.filter((student) => student.status !== 'removed').map((student) => student.id),
    )

    setSelectedStudentsForMassUpdate((previous) => {
      const next = new Set<string>()
      previous.forEach((id) => {
        if (activeIds.has(id)) {
          next.add(id)
        }
      })
      return next.size === previous.size ? previous : next
    })
  }, [selectedGroupMembers])

  const summarizedBalanceResults = useMemo(() => {
    if (!balanceResults) {
      return [] as BalanceChange[]
    }

    return summarizeBalanceChanges(balanceResults)
  }, [balanceResults])

  const balanceHistoryByStudent = useMemo(() => {
    const grouped = new Map<string, { studentId: string; studentName: string; classGroup: string; finalSubjectsSummary: Array<{ blockNumber: string; subjects: string }>; runs: Array<{ runId: string; createdAt: string; message: string; changes: BalanceChange[] }> }>()
    const studentClassById = new Map<string, string>()
    const studentById = new Map<string, StudentRecord>()
    ;(parsedData?.students || []).forEach((student) => {
      studentClassById.set(student.id, student.classGroup || '')
      studentById.set(student.id, student)
    })

    balanceHistory.forEach((run) => {
      const summarized = summarizeBalanceChanges(run.changes)
      if (summarized.length === 0) {
        return
      }

      const byStudent = new Map<string, BalanceChange[]>()
      summarized.forEach((change) => {
        if (!byStudent.has(change.studentId)) {
          byStudent.set(change.studentId, [])
        }
        byStudent.get(change.studentId)?.push(change)
      })

      byStudent.forEach((studentChanges, studentId) => {
        const studentName = studentChanges[0]?.studentName || `Student ${studentId}`
        const classGroup = studentClassById.get(studentId) || ''
        const finalSubjectsSummary = getFinalSubjectsByBlock(studentById.get(studentId))
        if (!grouped.has(studentId)) {
          grouped.set(studentId, { studentId, studentName, classGroup, finalSubjectsSummary, runs: [] })
        }

        grouped.get(studentId)?.runs.push({
          runId: run.id,
          createdAt: run.createdAt,
          message: run.message,
          changes: [...studentChanges].sort((a, b) => {
            const subjectNameCmp = (a.subjectName || a.subjectCode).localeCompare(b.subjectName || b.subjectCode)
            if (subjectNameCmp !== 0) {
              return subjectNameCmp
            }
            return a.subjectCode.localeCompare(b.subjectCode)
          }),
        })
      })
    })

    return Array.from(grouped.values()).sort((a, b) => a.studentName.localeCompare(b.studentName, 'nb-NO'))
  }, [balanceHistory, parsedData])

  const balanceSummaryHistory = useMemo(() => {
    return balanceHistory.filter((run) => run.message.trim().length > 0)
  }, [balanceHistory])

  const latestBalanceSummary = balanceSummaryHistory.length > 0
    ? balanceSummaryHistory[balanceSummaryHistory.length - 1]
    : null

  const allCollisionErrors = useMemo(() => {
    const errors: CollisionError[] = []
    const seenIds = new Set<string>()

    for (const run of [...balanceHistory].reverse()) {
      for (const err of run.collisionErrors ?? []) {
        if (!seenIds.has(err.studentId)) {
          seenIds.add(err.studentId)
          errors.push(err)
        }
      }
    }

    return errors
  }, [balanceHistory])

  const appendBalanceHistoryRun = (changes: BalanceChange[], message: string, collisionErrors?: CollisionError[]): void => {
    const summarized = summarizeBalanceChanges(changes)
    if (summarized.length === 0 && (!collisionErrors || collisionErrors.length === 0) && message.trim().length === 0) {
      return
    }

    const run: BalanceResultRun = {
      id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      createdAt: new Date().toISOString(),
      message,
      changes: summarized,
      ...(collisionErrors && collisionErrors.length > 0 ? { collisionErrors } : {}),
    }

    setBalanceHistory((previous) => [...previous, run])
  }

  const appendManualBalanceHistoryRun = (changes: BalanceChange[], message: string): void => {
    const summarized = summarizeBalanceChanges(changes)
    if (summarized.length === 0) {
      return
    }

    setBalanceHistory((previous) => {
      const nextHistory = previous.map((run) => ({ ...run, changes: [...run.changes] }))
      const remaining: BalanceChange[] = []

      summarized.forEach((change) => {
        let canceled = false

        for (let runIndex = nextHistory.length - 1; runIndex >= 0 && !canceled; runIndex -= 1) {
          const run = nextHistory[runIndex]
          for (let changeIndex = run.changes.length - 1; changeIndex >= 0; changeIndex -= 1) {
            const existing = run.changes[changeIndex]
            if (isInverseBalanceChange(existing, change)) {
              run.changes.splice(changeIndex, 1)
              canceled = true
              break
            }
          }
        }

        if (!canceled) {
          remaining.push(change)
        }
      })

      const cleanedHistory = nextHistory.filter((run) => run.changes.length > 0)
      if (remaining.length === 0) {
        return cleanedHistory
      }

      const run: BalanceResultRun = {
        id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
        createdAt: new Date().toISOString(),
        message,
        changes: remaining,
      }

      return [...cleanedHistory, run]
    })
  }

  const handleExportBalanceResultsWord = (): void => {
    if (balanceHistory.length === 0 && allCollisionErrors.length === 0) {
      setErrorMessage('Ingen balanseringsresultater a eksportere enda.')
      return
    }

    const now = new Date()
    const timestampForFile = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}-${String(now.getHours()).padStart(2, '0')}${String(now.getMinutes()).padStart(2, '0')}`
    const fileName = `balanseringsresultater-${timestampForFile}.doc`

    const collisionRows = allCollisionErrors.length > 0
      ? `<section><h2>Uløsbare blokk-kollisjoner</h2><p class="intro">Disse elevene har valgt fag som ikke kan plasseres i ulike blokker med gjeldende gruppetilbud. Manuell overstyring er nødvendig.</p>${allCollisionErrors
          .map((err) => {
            const subjectRows = err.subjects
              .map((subject) => `<li>${escapeHtml(formatSubjectDisplayName(subject.subjectName || subject.subjectCode))} (${escapeHtml(subject.block)})</li>`)
              .join('')

            return `<div class="collision-item"><div class="student-title"><strong>${escapeHtml(err.studentName)}</strong> (${escapeHtml(err.classGroup || '-')}) (${escapeHtml(err.studentId)})</div><ul>${subjectRows}</ul></div>`
          })
          .join('')}</section>`
      : ''

    const studentRows = balanceHistoryByStudent
      .map((student) => {
        const finalSubjectsSummary = student.finalSubjectsSummary
          .map((item) => `<span class="final-subject"><strong>${escapeHtml(item.blockNumber)}:</strong> ${escapeHtml(item.subjects)}</span>`)
          .join(' ')

        const netChanges = summarizeBalanceChanges(student.runs.flatMap((run) => run.changes))
          .sort((a, b) => {
            const subjectNameCmp = formatSubjectDisplayName(a.subjectName || a.subjectCode).localeCompare(
              formatSubjectDisplayName(b.subjectName || b.subjectCode),
              'nb-NO',
            )
            if (subjectNameCmp !== 0) {
              return subjectNameCmp
            }
            return a.subjectCode.localeCompare(b.subjectCode)
          })

        const subjectHtml = netChanges
          .map((change) => {
            const isAdded = !change.fromGroupCode.trim() && !change.fromBlock.trim() && (change.toGroupCode.trim() || change.toBlock.trim())
            const isRemoved = !change.toGroupCode.trim() && !change.toBlock.trim() && (change.fromGroupCode.trim() || change.fromBlock.trim())
            const lineClass = isAdded ? 'change-added' : isRemoved ? 'change-removed' : 'change-moved'

            return `<div class="change-line ${lineClass}"><span class="change-main"><strong>${escapeHtml(formatSubjectDisplayName(change.subjectName || change.subjectCode))}</strong>: ${escapeHtml(formatBalanceChangeText(change))}</span></div>`
          })
          .join('')

        return `<section class="student"><h3>${escapeHtml(student.studentName)} (${escapeHtml(student.classGroup || '-')}) (${escapeHtml(student.studentId)})</h3>${finalSubjectsSummary ? `<div class="student-final-subjects">${finalSubjectsSummary}</div>` : ''}<div class="student-change-block">${subjectHtml}</div><div class="student-spacer">&nbsp;</div></section>`
      })
      .join('')

    const summaryHistoryRows = balanceSummaryHistory.length > 0
      ? `<section><h2>Historikk</h2>${[...balanceSummaryHistory].reverse()
          .map((run) => `<div class="summary-history-item"><span>${escapeHtml(run.message)}</span><span>${escapeHtml(formatTimestamp(run.createdAt))}</span></div>`)
          .join('')}</section>`
      : ''

    const htmlDocument = `<!doctype html><html><head><meta charset="utf-8"><title>Balanseringsresultater</title><style>body{font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#1f2b3d;margin:24px;}h1{font-size:18pt;margin:0 0 14px;}h2{font-size:13pt;margin:18px 0 8px;padding-bottom:4px;border-bottom:1px solid #d9e3f0;}h3{font-size:13pt;margin:0 0 6px;}.intro{margin:0 0 10px;color:#5b6d86;}.student{padding:8px 0;}.collision-item{margin-bottom:10px;padding:8px 10px;border:1px solid #f0c8cc;border-radius:8px;background:#fff5f5;}.student-title{margin-bottom:4px;}.student-final-subjects{margin:0 0 12px 0;color:#425775;}.student-change-block{margin-top:12px;}.student-spacer{height:24pt;line-height:24pt;font-size:1pt;}.change-line{margin-top:5px;padding:4px 7px;border-left:3px solid transparent;border-radius:6px;}.change-main{min-width:0;}.change-moved{background:#eef5ff;border-left-color:#2a63b7;color:#1f3f6c;}.change-added{background:#eaf9f0;border-left-color:#2f8f5b;color:#1f5d3d;}.change-removed{background:#fff1f1;border-left-color:#c45555;color:#7a2c2c;}.final-subject{margin-right:8px;}.summary-history-item{display:flex;justify-content:space-between;gap:10px;padding:6px 8px;border-top:1px solid #eef3f9;color:#39506f;}.summary-history-item:first-of-type{border-top:none;}.summary-history-item span:last-child{flex-shrink:0;color:#687c98;font-size:8.5pt;white-space:nowrap;}ul{margin:4px 0 0;padding-left:18px;}li{margin:2px 0;}</style></head><body><h1>Balanseringsresultater</h1>${collisionRows}${studentRows ? `<section><h2>Elevendringer</h2>${studentRows}</section>` : ''}${summaryHistoryRows}</body></html>`

    const blob = new Blob(['\ufeff', htmlDocument], { type: 'application/msword;charset=utf-8' })
    const url = URL.createObjectURL(blob)
    const anchor = document.createElement('a')
    anchor.href = url
    anchor.download = fileName
    document.body.appendChild(anchor)
    anchor.click()
    anchor.remove()
    URL.revokeObjectURL(url)
    setErrorMessage('')
  }

  const handleExport = (): void => {
    if (!parsedData || !isParsedDataExportReady(parsedData)) {
      setErrorMessage('Ingen gyldig eksportkilde er lastet. Last opp filen pa nytt for eksport.')
      return
    }

    try {
      const exportText = buildExportText(parsedData)
      const now = new Date()
      const yyyymmdd = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`
      const downloadName = `${yyyymmdd}timeplan.txt`

      const encodedExportBuffer = encodeWindows1252(exportText)
      const blob = new Blob([encodedExportBuffer], { type: 'text/plain;charset=windows-1252' })
      const url = URL.createObjectURL(blob)
      const anchor = document.createElement('a')
      anchor.href = url
      anchor.download = downloadName
      document.body.appendChild(anchor)
      anchor.click()
      anchor.remove()
      URL.revokeObjectURL(url)
      setErrorMessage('')
    } catch {
      setErrorMessage('Eksport feilet. Last opp filen pa nytt og forsok igjen.')
    }
  }

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
    const file = event.target.files?.[0]
    if (!file) {
      return
    }

    try {
      const fileBuffer = await file.arrayBuffer()
      const text = decodeBestEffort(fileBuffer)
      const parsed = parseNovaschemExport(text, file.name)
      setUndoStack([])
      setRedoStack([])
      lastHistorySnapshotRef.current = null
      isHistoryNavigationRef.current = false
      setParsedData(parsed)
      setSelectedStudentId(parsed.students[0]?.id || '')
      setSelectedGroupKey('')
      setStudentQuery('')
      setSubjectQuery('')
      setBlockFilter('')
      setErrorMessage('')
    } catch {
      setUndoStack([])
      setRedoStack([])
      lastHistorySnapshotRef.current = null
      isHistoryNavigationRef.current = false
      setParsedData(null)
      setSelectedStudentId('')
      setSelectedGroupKey('')
      setErrorMessage('Kunne ikke lese eksportfilen. Bekreft at filen er en Novaschem TXT-eksport.')
    } finally {
      // Reset input so choosing the same file again triggers onChange.
      event.target.value = ''
    }
  }

  return (
    <div className="app-shell">
      <header className="hero">
        <p className="hero-tag">Novaschem til SATS</p>
        <h1>Novaschem Eksportvisning</h1>
        <p className="hero-subtitle">
          Last opp en Novaschem TXT-eksport for å bla gjennom elever, deres valgte fag og tildelte blokker.
        </p>
      </header>

      <section className="upload-panel">
        <label htmlFor="novaschem-file" className="file-label">
          Velg Novaschem TXT-fil
        </label>
        <div className="upload-actions-row">
          <input id="novaschem-file" ref={fileInputRef} type="file" accept=".txt" onChange={handleFileUpload} />
          <div className="storage-controls">
          <button type="button" onClick={clearStoredData} className="storage-button danger">
            Tøm lagret data
          </button>
          </div>
        </div>
        {errorMessage && <p className="error-text">{errorMessage}</p>}
        {parsedData && (
          <div className="stats-grid">
            <article>
              <strong>{parsedData.students.length}</strong>
              <span>Elever</span>
            </article>
            <article>
              <strong>{parsedData.subjects.length}</strong>
              <span>Fag i bruk</span>
            </article>
            <article>
              <strong>{parsedData.blocks.length}</strong>
              <span>Blokker funnet</span>
            </article>
            <article>
              <strong>{parsedData.tableNames.length}</strong>
              <span>Tabeller lest</span>
            </article>
          </div>
        )}
      </section>

      {parsedData && (
        <>
          <section className="controls-row">
            <div className="switch-group">
              <button
                type="button"
                className={viewMode === 'students' ? 'active' : ''}
                onClick={() => setViewMode('students')}
              >
                Elevvisning
              </button>
              <button
                type="button"
                className={viewMode === 'subjects' ? 'active' : ''}
                onClick={() => setViewMode('subjects')}
              >
                Fagvisning
              </button>
              <button
                type="button"
                className={viewMode === 'blokkoversikt' ? 'active' : ''}
                onClick={() => setViewMode('blokkoversikt')}
              >
                Blokkoversikt
              </button>
              <button
                type="button"
                className={viewMode === 'bytteoversikt' ? 'active' : ''}
                onClick={() => setViewMode('bytteoversikt')}
              >
                Blokkvalg
              </button>
            </div>
            <div className="controls-actions">
              <button type="button" className="clear-results-button history-button" onClick={handleUndo} disabled={undoStack.length === 0}>
                Undo ({undoStack.length})
              </button>
              <button type="button" className="clear-results-button history-button" onClick={handleRedo} disabled={redoStack.length === 0}>
                Redo ({redoStack.length})
              </button>
              <button
                type="button"
                onClick={() => setShowProgressiveBalanceDialog(true)}
                className="balance-button"
              >
                Progressiv balansering
              </button>
              <button type="button" className="export-button" onClick={handleExport}>
                Eksporter TXT
              </button>
            </div>
          </section>

          <section className="secondary-controls-row">
            <input
              type="search"
              value={subjectQuery}
              onChange={(event) => setSubjectQuery(event.target.value)}
              placeholder="Filtrer etter fagkode eller navn"
            />

            <label className="blokkfag-checkbox">
              <input
                type="checkbox"
                checked={onlyBlokkfag}
                onChange={(event) => setOnlyBlokkfag(event.target.checked)}
              />
              <span>Kun blokkfag</span>
            </label>

            <select value={blockFilter} onChange={(event) => setBlockFilter(event.target.value)}>
              <option value="">Alle blokker</option>
              {parsedData.blocks.map((block) => (
                <option key={block} value={block}>
                  {block}
                </option>
              ))}
            </select>

            {viewMode === 'students' && (
              <div className="student-filter-actions">
                <button
                  type="button"
                  disabled={studentFilterCounts.missingSubjectsCount === 0}
                  onClick={() => {
                    const next = !showIncompleteBlocks
                    setShowIncompleteBlocks(next)
                    if (next) { setShowOverloadedStudents(false); setShowBlockCollisions(false); setShowDuplicateSubjects(false) }
                  }}
                  className={`filter-button ${showIncompleteBlocks ? 'active' : ''}`}
                >
                  Mangler fag ({studentFilterCounts.missingSubjectsCount})
                </button>
                <button
                  type="button"
                  disabled={studentFilterCounts.tooManySubjectsCount === 0}
                  onClick={() => {
                    const next = !showOverloadedStudents
                    setShowOverloadedStudents(next)
                    if (next) { setShowIncompleteBlocks(false); setShowBlockCollisions(false); setShowDuplicateSubjects(false) }
                  }}
                  className={`filter-button ${showOverloadedStudents ? 'active' : ''}`}
                >
                  For mange fag ({studentFilterCounts.tooManySubjectsCount})
                </button>
                <button
                  type="button"
                  disabled={studentFilterCounts.blockCollisionsCount === 0}
                  onClick={() => {
                    const next = !showBlockCollisions
                    setShowBlockCollisions(next)
                    if (next) { setShowIncompleteBlocks(false); setShowOverloadedStudents(false); setShowDuplicateSubjects(false) }
                  }}
                  className={`filter-button ${showBlockCollisions ? 'active' : ''}`}
                >
                  Blokk-kollisjoner ({studentFilterCounts.blockCollisionsCount})
                </button>
                <button
                  type="button"
                  disabled={studentFilterCounts.duplicateSubjectsCount === 0}
                  onClick={() => {
                    const next = !showDuplicateSubjects
                    setShowDuplicateSubjects(next)
                    if (next) { setShowIncompleteBlocks(false); setShowOverloadedStudents(false); setShowBlockCollisions(false) }
                  }}
                  className={`filter-button ${showDuplicateSubjects ? 'active' : ''}`}
                >
                  Duplikater ({studentFilterCounts.duplicateSubjectsCount})
                </button>
              </div>
            )}
          </section>

          {viewMode === 'students' ? (
            <section className="viewer-grid">
              <aside className="student-list-panel">
                <input
                  type="search"
                  value={studentQuery}
                  onChange={(event) => setStudentQuery(event.target.value)}
                  placeholder="Søk elevnavn eller nummer"
                />

                <div className="student-list">
                  {filteredStudents.map((student) => (
                    <button
                      key={student.id}
                      type="button"
                      onClick={() => setSelectedStudentId(student.id)}
                      className={student.id === selectedStudentId ? 'student-row active' : 'student-row'}
                    >
                      <span>
                        {student.fullName} {student.classGroup && `(${student.classGroup})`}
                      </span>
                      <small>
                        {student.id} | {student.assignments.length} fag
                      </small>
                    </button>
                  ))}
                </div>
              </aside>

              <article className="detail-panel" onClick={() => setPendingRemovalAssignment('')}>
                {selectedStudent ? (
                  <>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                      <div>
                        <h2>
                          {selectedStudent.fullName} {selectedStudent.classGroup && `(${selectedStudent.classGroup})`}
                        </h2>
                        <p>
                          Elevnummer: <strong>{selectedStudent.id}</strong>
                          {selectedStudent.email ? <> | {selectedStudent.email}</> : null}
                        </p>
                      </div>
                      <button
                        type="button"
                        onClick={(e) => {
                          e.stopPropagation()
                          setShowStudentAddSubjectDialog(true)
                        }}
                        className="balance-button"
                      >
                        Legg til fag
                      </button>
                    </div>

                    <table>
                      <thead>
                        <tr>
                          <th>Fag</th>
                          <th>Tittel</th>
                          <th>Gruppe</th>
                          <th>Blokk</th>
                          <th style={{ width: '170px' }}>Handling</th>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredAssignments.map((assignment) => {
                          const assignmentKey = `${assignment.subjectCode}|${assignment.groupCode}|${assignment.block}`
                          const isPendingRemoval = pendingRemovalAssignment === assignmentKey

                          return (
                            <tr
                              key={`${assignment.subjectCode}-${assignment.groupCode}-${assignment.block}`}
                              className={assignment.block === 'MATTE' ? 'matte-assignment-row' : ''}
                            >
                              <td>{assignment.subjectCode}</td>
                              <td>{assignment.subjectName}</td>
                              <td>{assignment.groupCode}</td>
                              <td>{assignment.block || '-'}</td>
                              <td>
                                <button
                                  type="button"
                                  onClick={(e) => {
                                    e.stopPropagation()
                                    setPendingRemovalAssignment('')
                                    setStudentSwapAssignmentKey(assignmentKey)
                                    setStudentSwapTargetSubject(assignment.subjectCode)
                                    setStudentSwapTargetBlock('')
                                    setShowStudentSwapDialog(true)
                                  }}
                                  style={{
                                    padding: '0.25rem 0.5rem',
                                    fontSize: '0.875rem',
                                    backgroundColor: '#0969da',
                                    color: '#ffffff',
                                    border: '1px solid #0969da',
                                    borderRadius: '4px',
                                    cursor: 'pointer',
                                    marginRight: '0.35rem',
                                  }}
                                >
                                  Bytt
                                </button>
                                <button
                                  type="button"
                                  onClick={(e) => {
                                    e.stopPropagation()

                                    if (!isPendingRemoval) {
                                      // First click: Enter confirmation mode
                                      setPendingRemovalAssignment(assignmentKey)
                                      return
                                    }

                                    // Second click: Actually remove
                                    const updatedStudents = parsedData.students.map((student) => {
                                      if (student.id !== selectedStudent.id) {
                                        return student
                                      }

                                      return {
                                        ...student,
                                        assignments: student.assignments.filter(
                                          (a) => !(a.subjectCode === assignment.subjectCode && a.groupCode === assignment.groupCode && a.block === assignment.block)
                                        ),
                                      }
                                    })

                                    // Recalculate breakdowns
                                    const { groupBreakdowns, blockBreakdowns } = recalculateBreakdowns(updatedStudents, parsedData.groupBreakdowns)

                                    // Update subjects with new student count
                                    const updatedSubjects = parsedData.subjects.map((subject) => {
                                      if (subject.code !== assignment.subjectCode) {
                                        return subject
                                      }
                                      return {
                                        ...subject,
                                        studentCount: Math.max(0, subject.studentCount - 1),
                                      }
                                    })

                                    setParsedData({
                                      ...parsedData,
                                      students: updatedStudents,
                                      subjects: updatedSubjects,
                                      groupBreakdowns,
                                      blockBreakdowns,
                                    })

                                    appendBalanceHistoryRun(
                                      [
                                        {
                                          studentId: selectedStudent.id,
                                          studentName: selectedStudent.fullName,
                                          subjectCode: assignment.subjectCode,
                                          subjectName: assignment.subjectName,
                                          fromGroupCode: assignment.groupCode,
                                          fromBlock: assignment.block,
                                          toGroupCode: '',
                                          toBlock: '',
                                        },
                                      ],
                                      `${assignment.subjectName} fjernet for ${selectedStudent.fullName}`,
                                    )

                                    setPendingRemovalAssignment('')
                                    setBalanceMessage(`${assignment.subjectName} fjernet for ${selectedStudent.fullName}`)
                                  }}
                                  style={{
                                    padding: '0.25rem 0.5rem',
                                    fontSize: '0.875rem',
                                    backgroundColor: isPendingRemoval ? '#ffc107' : '#dc3545',
                                    color: '#ffffff',
                                    border: isPendingRemoval ? '1px solid #ffc107' : '1px solid #dc3545',
                                    borderRadius: '4px',
                                    cursor: 'pointer',
                                  }}
                                >
                                  {isPendingRemoval ? 'Bekreft' : 'Fjern'}
                                </button>
                              </td>
                            </tr>
                          )
                        })}
                      </tbody>
                    </table>
                  </>
                ) : (
                  <p>Velg en elev for å se fagvalg.</p>
                )}
              </article>
            </section>
          ) : viewMode === 'subjects' ? (
            <section className="subject-panel">
              <div className="subject-view-header">
                <h2>Fagvisning</h2>
              </div>

              {(balanceResults !== null || balanceHistory.length > 0) && (
                <div className="balance-results">
                  <h3>
                    Balanseringsresultater{' '}
                    {balanceHistoryByStudent.length > 0
                      ? `(${balanceHistoryByStudent.length} elev${balanceHistoryByStudent.length === 1 ? '' : 'er'} med historikk)`
                      : ''}
                  </h3>
                  {latestBalanceSummary && (
                    <div className="balance-message-panel">
                      <div className="balance-message">{latestBalanceSummary.message}</div>
                      {balanceSummaryHistory.length > 1 && (
                        <button
                          type="button"
                          onClick={() => setShowBalanceMessageHistory((previous) => !previous)}
                          className="balance-message-toggle"
                        >
                          {showBalanceMessageHistory ? 'Skjul historikk' : 'Vis historikk'} ({balanceSummaryHistory.length})
                        </button>
                      )}
                      {showBalanceMessageHistory && (
                        <div className="balance-message-history">
                          {[...balanceSummaryHistory].reverse().map((run) => (
                            <div key={run.id} className="balance-message-history-item">
                              <span>{run.message}</span>
                              <span>{formatTimestamp(run.createdAt)}</span>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  )}
                  {balanceHistoryByStudent.length > 0 && (
                    <div className="balance-list">
                      {balanceHistoryByStudent.map((student) => (
                        <div key={student.studentId} className="balance-item">
                          <strong>{student.studentName}</strong> ({student.classGroup || '-'}) ({student.studentId})
                          {student.finalSubjectsSummary.length > 0 && (
                            <>
                              {' - '}
                              {student.finalSubjectsSummary.map((item, index) => (
                                <span key={`${student.studentId}-${item.blockNumber}`}>
                                  {index > 0 ? ' ' : ''}
                                  <strong>{item.blockNumber}:</strong> {item.subjects}
                                </span>
                              ))}
                            </>
                          )}
                          {groupBalanceRunsBySubject(student.runs).map((subjectGroup) => (
                            <div key={`${student.studentId}-${subjectGroup.subjectCode}`} className="balance-subject-group">
                              {subjectGroup.changes.map(({ runId, createdAt, message, change }) => (
                                <div
                                  key={`${runId}-${change.subjectCode}-${change.fromGroupCode}-${change.toGroupCode}-${change.fromBlock}-${change.toBlock}`}
                                  className={`balance-change-line ${
                                    !change.fromGroupCode.trim() && !change.fromBlock.trim() && (change.toGroupCode.trim() || change.toBlock.trim())
                                      ? 'balance-change-added'
                                      : !change.toGroupCode.trim() && !change.toBlock.trim() && (change.fromGroupCode.trim() || change.fromBlock.trim())
                                        ? 'balance-change-removed'
                                        : 'balance-change-moved'
                                  }`}
                                >
                                  <span className="balance-change-main">
                                    <strong>{subjectGroup.subjectName}</strong>: {formatBalanceChangeText(change)}
                                  </span>
                                  <span className="balance-change-meta">{getBalanceRunLabel(message || 'Balansering')} {formatTimestamp(createdAt)}</span>
                                </div>
                              ))}
                            </div>
                          ))}
                        </div>
                      ))}
                    </div>
                  )}
                  {balanceResults !== null && balanceResults.length > 0 && summarizedBalanceResults.length === 0 && (
                    <p style={{ fontSize: '0.9rem', color: '#666', marginBottom: '0.8rem' }}>
                      Ingen netto fagendringer etter oppsummering av mellomsteg.
                    </p>
                  )}
                  {allCollisionErrors.length > 0 && (
                    <div className="collision-error-section">
                      <strong className="collision-error-heading">
                        ⚠ Uløsbare blokk-kollisjoner ({allCollisionErrors.length} elev{allCollisionErrors.length === 1 ? '' : 'er'})
                      </strong>
                      <p className="collision-error-desc">
                        Disse elevene har valgt fag som ikke kan plasseres i ulike blokker med gjeldende gruppetilbud. Manuell overstyring er nødvendig.
                      </p>
                      <div className="balance-list">
                        {allCollisionErrors.map((err) => (
                          <div key={err.studentId} className="balance-item collision-error-item">
                            <strong>{err.studentName}</strong> ({err.classGroup || '-'}) — {err.studentId}
                            <ul className="collision-error-subjects">
                              {err.subjects.map((s) => (
                                <li key={`${s.subjectCode}-${s.block}`}>
                                  {formatSubjectDisplayName(s.subjectName || s.subjectCode)} ({s.block})
                                </li>
                              ))}
                            </ul>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}
                  {balanceResults !== null && balanceResults.length === 0 && debugGroups.length > 0 && (
                    <>
                      <p style={{ fontSize: '0.9rem', color: '#666', marginBottom: '0.8rem' }}>
                        Alle grupper analysert ({debugGroups.length} grupper):
                      </p>
                      <div className="balance-list" style={{ maxHeight: '400px' }}>
                        {debugGroups.map((group, idx) => (
                          <div key={idx} className="balance-item" style={{ fontSize: '0.85rem' }}>
                            <strong>{group.key}</strong>
                            <br />
                            Antall: {group.count}, Maks: {group.maxCap}
                            <br />
                            <span style={{ color: group.status.includes('OVERFULL') ? '#c41e3a' : '#2f7044' }}>
                              {group.status}
                            </span>
                          </div>
                        ))}
                      </div>
                    </>
                  )}
                  <button
                    type="button"
                    onClick={handleExportBalanceResultsWord}
                    className="clear-results-button"
                    disabled={balanceHistoryByStudent.length === 0}
                  >
                    Eksporter resultat til Word
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      setBalanceResults(null)
                      setBalanceHistory([])
                      setBalanceMessage('')
                      setDebugGroups([])
                    }}
                    className="clear-results-button"
                  >
                    Tøm
                  </button>
                </div>
              )}

              <div className="subject-actions-under-results">
                <button
                  type="button"
                  onClick={() => setShowAddSubjectDialog(true)}
                  className="balance-button"
                >
                  Legg til fag
                </button>
              </div>

              <h2>Per blokk</h2>
              <div className="block-summary-grid">
                {parsedData.blockBreakdowns
                  .filter((item) => {
                    const isStandardBlock = /^Blokk \d+$/u.test(item.block)
                    return blockFilter ? item.block === blockFilter : isStandardBlock
                  })
                  .map((item) => {
                    const blockDelta = balanceBlockDeltaCounts.get(item.block)
                    return (
                      <article key={item.block}>
                        <strong>{item.block}</strong>
                        <span>{item.subjectCount} fag</span>
                        <span>{item.groupCount} grupper</span>
                        <span>
                          {item.studentCount} elever
                          {blockDelta !== undefined && (
                            <span style={{ marginLeft: '0.5rem', color: blockDelta > 0 ? '#2f7044' : '#c41e3a', fontWeight: 'bold' }}>
                              {blockDelta > 0 ? '+' : ''}{blockDelta}
                            </span>
                          )}
                        </span>
                      </article>
                    )
                  })}
              </div>

              <h2>Per fag (alle grupper)</h2>
              <table>
                <thead>
                  <tr>
                    <th>Fag</th>
                    <th>Tittel</th>
                    {perSubjectBlockColumns.map((block) => (
                      <th key={`head-${block}`}>{block}</th>
                    ))}
                    <th>Total</th>
                  </tr>
                </thead>
                <tbody>
                  {perSubjectMatrixRows.filter((row) => row.total > 0).map((row) => (
                    <tr key={row.subject.code} className={row.isOverfilled ? 'subject-overfilled-row' : ''}>
                      <td>{row.subject.code}</td>
                      <td>{row.subject.name}</td>
                      {perSubjectBlockColumns.map((block) => {
                        const hasGroup = row.hasGroupByBlock.get(block) || false
                        const count = row.countsByBlock.get(block) || 0
                        const isOverLimit = hasGroup && count > row.maxCap

                        return (
                          <td key={`${row.subject.code}-${block}`} className={isOverLimit ? 'subject-over-limit' : ''}>
                            {hasGroup ? count : '-'}
                          </td>
                        )
                      })}
                      <td>{row.total}</td>
                    </tr>
                  ))}
                </tbody>
              </table>

              <h2>Per faggruppe</h2>
              <div className="per-faggruppe-toolbar">
                <label htmlFor="per-faggruppe-sort">Sortér:</label>
                <select
                  id="per-faggruppe-sort"
                  value={perFaggruppeSortBy}
                  onChange={(event) => setPerFaggruppeSortBy(event.target.value as 'blokk' | 'tittel' | 'students' | 'change')}
                >
                  <option value="blokk">Blokk</option>
                  <option value="tittel">Tittel</option>
                  <option value="students">Antall elever</option>
                  <option value="change">Endring</option>
                </select>
              </div>
              <table>
                <thead>
                  <tr>
                    <th>Fag</th>
                    <th>Tittel</th>
                    <th>Gruppe</th>
                    <th>Elever</th>
                    <th>Blokk</th>
                    {balanceDeltaCounts.size > 0 && <th>Endring</th>}
                  </tr>
                </thead>
                <tbody>
                  {filteredGroupBreakdowns.filter((item) => item.studentCount > 0).map((item) => {
                    const groupKey = `${item.subjectCode}|${item.groupCode}|${item.block}`
                    const delta = balanceDeltaCounts.get(groupKey)
                    const isExpanded = selectedGroupKey === `${item.subjectCode}-${item.groupCode}-${item.block}`
                    const colSpan = balanceDeltaCounts.size > 0 ? 6 : 5

                    return (
                      <>
                        <tr
                          key={`${item.subjectCode}-${item.groupCode}-${item.block}`}
                          onClick={() => {
                            const key = `${item.subjectCode}-${item.groupCode}-${item.block}`
                            setSelectedGroupKey(isExpanded ? '' : key)
                          }}
                          className={isExpanded ? 'clickable-row active' : 'clickable-row'}
                        >
                          <td>{item.subjectCode}</td>
                          <td>{item.subjectName}</td>
                          <td>{item.groupCode}</td>
                          <td>{item.studentCount}</td>
                          <td>{item.block || '-'}</td>
                          {balanceDeltaCounts.size > 0 && (
                            <td>
                              {delta !== undefined ? (
                                <span style={{ color: delta > 0 ? '#2f7044' : '#c41e3a', fontWeight: 'bold' }}>
                                  {delta > 0 ? '+' : ''}{delta}
                                </span>
                              ) : (
                                <span style={{ color: '#999' }}>-</span>
                              )}
                            </td>
                          )}
                        </tr>
                        {isExpanded && (
                          <tr key={`${item.subjectCode}-${item.groupCode}-${item.block}-details`} className="expanded-row">
                            <td colSpan={colSpan} style={{ padding: '1rem', backgroundColor: '#f8f9fa' }}>
                              <div style={{ marginBottom: '1rem' }}>
                                <strong>Elever i valgt faggruppe</strong>
                                <p style={{ margin: '0.25rem 0 0.75rem 0', fontSize: '0.9rem', color: '#666' }}>
                                  {selectedGroupActiveMemberCount} aktive i gruppen.
                                  {selectedGroupMembers.some((student) => student.status === 'added') && ' Nye elever er markert i grønt.'}
                                  {selectedGroupMembers.some((student) => student.status === 'removed') && ' Fjernede elever vises nederst i rødt og telles ikke med.'}
                                </p>
                              </div>
                              <div style={{ marginBottom: '1rem', display: 'flex', gap: '0.5rem' }}>
                                <button
                                  type="button"
                                  onClick={(e) => {
                                    e.stopPropagation()
                                    setPendingMassRemoval(false)
                                    const [currentSubjectCode] = selectedGroupKey.split('-')
                                    setMassUpdateTargetSubject(currentSubjectCode || '')
                                    setMassUpdateTargetBlock('')
                                    setShowMassUpdateDialog(true)
                                  }}
                                  disabled={selectedStudentsForMassUpdate.size === 0}
                                  className="balance-button"
                                >
                                  Massoppdater ({selectedStudentsForMassUpdate.size} valgt)
                                </button>
                                <button
                                  type="button"
                                  onClick={(e) => {
                                    e.stopPropagation()
                                    setSelectedStudentsForMassUpdate(new Set())
                                  }}
                                  disabled={selectedStudentsForMassUpdate.size === 0}
                                  className="clear-results-button"
                                >
                                  Tøm valg
                                </button>
                              </div>
                              <div className="group-member-list">
                                {selectedGroupMembers.map((student) => (
                                  <div
                                    key={`${student.status}-${student.id}`}
                                    className={`group-member-item ${
                                      selectedStudentsForMassUpdate.has(student.id) ? 'group-member-selected' : ''
                                    } ${
                                      student.status === 'added'
                                        ? 'group-member-added'
                                        : student.status === 'removed'
                                          ? 'group-member-removed'
                                          : ''
                                    }`}
                                    style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}
                                    onClick={(e) => {
                                      e.stopPropagation()
                                      if (student.status === 'removed') {
                                        return
                                      }

                                      const next = new Set(selectedStudentsForMassUpdate)
                                      if (next.has(student.id)) {
                                        next.delete(student.id)
                                      } else {
                                        next.add(student.id)
                                      }
                                      setSelectedStudentsForMassUpdate(next)
                                    }}
                                  >
                                    <div style={{ flex: 1 }}>
                                      <span>{student.fullName}</span>{' '}
                                      <small>
                                        {student.classGroup ? `${student.classGroup} | ` : ''}
                                        {student.id}
                                        {student.status === 'added' ? ' | Lagt til' : ''}
                                        {student.status === 'removed' ? ' | Fjernet' : ''}
                                      </small>
                                    </div>
                                  </div>
                                ))}
                              </div>
                            </td>
                          </tr>
                        )}
                      </>
                    )
                  })}
                </tbody>
              </table>

              {showMassUpdateDialog && (
                <div
                  className="modal-overlay"
                  onClick={() => {
                    setShowMassUpdateDialog(false)
                    setPendingMassRemoval(false)
                  }}
                >
                  <div className="modal-content" onClick={(e) => e.stopPropagation()}>
                    <h3>Massoppdater elever</h3>
                    <p>{selectedStudentsForMassUpdate.size} elever valgt</p>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '1rem', marginTop: '1rem' }}>
                      <label>
                        <strong>Nytt fag:</strong>
                        <select
                          value={massUpdateTargetSubject}
                          onChange={(e) => {
                            setMassUpdateTargetSubject(e.target.value)
                            setMassUpdateTargetBlock('')
                          }}
                          style={{ width: '100%', marginTop: '0.5rem', padding: '0.5rem' }}
                        >
                          <option value="">Velg fag...</option>
                          {parsedData?.subjects
                            .filter((s) => s.blocks && s.blocks.length > 0)
                            .sort((a, b) => a.name.localeCompare(b.name))
                            .map((subject) => (
                              <option key={subject.code} value={subject.code}>
                                {subject.code} - {subject.name}
                              </option>
                            ))}
                        </select>
                      </label>
                      <label>
                        <strong>Ny blokk:</strong>
                        <select
                          value={massUpdateTargetBlock}
                          onChange={(e) => setMassUpdateTargetBlock(e.target.value)}
                          style={{ width: '100%', marginTop: '0.5rem', padding: '0.5rem' }}
                        >
                          <option value="">Velg blokk...</option>
                          {(massUpdateTargetSubject
                            ? getSubjectBlockChoices(parsedData, massUpdateTargetSubject)
                            : []
                          ).map((choice) => (
                            <option key={choice.block} value={choice.block}>
                              {choice.block} ({choice.studentCount} elever)
                            </option>
                          ))}
                        </select>
                      </label>
                      <div style={{ display: 'flex', gap: '0.5rem', marginTop: '0.5rem' }}>
                        <button
                          type="button"
                          onClick={() => {
                            if (!massUpdateTargetSubject || !massUpdateTargetBlock || !selectedGroupKey) {
                              alert('Vennligst velg både fag og blokk')
                              return
                            }

                            // Parse current group key
                            const [oldSubjectCode, oldGroupCode, oldBlock] = selectedGroupKey.split('-')

                            // Find the target subject details
                            const targetSubject = parsedData?.subjects.find((s) => s.code === massUpdateTargetSubject)
                            if (!targetSubject) {
                              alert('Fag ikke funnet')
                              return
                            }

                            // Find the smallest group for the target subject/block combination
                            const targetGroups = parsedData?.groupBreakdowns.filter(
                              (gb) => gb.subjectCode === massUpdateTargetSubject && gb.block === massUpdateTargetBlock
                            ) || []

                            if (targetGroups.length === 0) {
                              alert(`Ingen grupper funnet for ${targetSubject.name} i ${massUpdateTargetBlock}`)
                              return
                            }

                            targetGroups.sort((a, b) => a.studentCount - b.studentCount)
                            const smallestGroup = targetGroups[0]
                            const manualChanges: BalanceChange[] = []

                            // Apply mass update by directly manipulating student assignments
                            const updatedStudents = parsedData!.students.map((student) => {
                              if (!selectedStudentsForMassUpdate.has(student.id)) {
                                return student
                              }

                              const oldAssignment = student.assignments.find(
                                (a) => a.subjectCode === oldSubjectCode && a.groupCode === oldGroupCode && a.block === oldBlock,
                              )
                              if (!oldAssignment) {
                                return student
                              }

                              if (oldAssignment.subjectCode === massUpdateTargetSubject) {
                                manualChanges.push({
                                  studentId: student.id,
                                  studentName: student.fullName,
                                  subjectCode: massUpdateTargetSubject,
                                  subjectName: targetSubject.name,
                                  fromGroupCode: oldAssignment.groupCode,
                                  fromBlock: oldAssignment.block,
                                  toGroupCode: smallestGroup.groupCode,
                                  toBlock: massUpdateTargetBlock,
                                })
                              } else {
                                manualChanges.push(
                                  {
                                    studentId: student.id,
                                    studentName: student.fullName,
                                    subjectCode: oldAssignment.subjectCode,
                                    subjectName: oldAssignment.subjectName,
                                    fromGroupCode: oldAssignment.groupCode,
                                    fromBlock: oldAssignment.block,
                                    toGroupCode: '',
                                    toBlock: '',
                                  },
                                  {
                                    studentId: student.id,
                                    studentName: student.fullName,
                                    subjectCode: massUpdateTargetSubject,
                                    subjectName: targetSubject.name,
                                    fromGroupCode: '',
                                    fromBlock: '',
                                    toGroupCode: smallestGroup.groupCode,
                                    toBlock: massUpdateTargetBlock,
                                  },
                                )
                              }

                              // Remove old assignment and add new one
                              const updatedAssignments = student.assignments
                                .filter((a) => !(a.subjectCode === oldSubjectCode && a.groupCode === oldGroupCode && a.block === oldBlock))
                                .concat([
                                  {
                                    subjectCode: massUpdateTargetSubject,
                                    subjectName: targetSubject.name,
                                    groupCode: smallestGroup.groupCode,
                                    block: massUpdateTargetBlock,
                                  },
                                ])

                              return {
                                ...student,
                                assignments: updatedAssignments,
                              }
                            })
                            const { groupBreakdowns: newGroupBreakdowns, blockBreakdowns: newBlockBreakdowns } = recalculateBreakdowns(updatedStudents, parsedData.groupBreakdowns)
                            const updatedData = {
                              ...parsedData!,
                              students: updatedStudents,
                              groupBreakdowns: newGroupBreakdowns,
                              blockBreakdowns: newBlockBreakdowns,
                            }
                            setParsedData(updatedData)

                            appendBalanceHistoryRun(
                              manualChanges,
                              `Masseoppdatering: ${manualChanges.length} elever flyttet til ${targetSubject.name} (${smallestGroup.groupCode}) i ${massUpdateTargetBlock}`,
                            )

                            // Update message
                            setBalanceMessage(`Masseoppdatering: ${selectedStudentsForMassUpdate.size} elever flyttet til ${targetSubject.name} (${smallestGroup.groupCode}) i ${massUpdateTargetBlock}`)

                            // Clear selection and close dialog
                            setSelectedStudentsForMassUpdate(new Set())
                            setShowMassUpdateDialog(false)
                            setPendingMassRemoval(false)
                            setMassUpdateTargetSubject('')
                            setMassUpdateTargetBlock('')
                            setSelectedGroupKey('')
                          }}
                          className="balance-button"
                          disabled={!massUpdateTargetSubject || !massUpdateTargetBlock}
                        >
                          Oppdater
                        </button>
                        <button
                          type="button"
                          onClick={() => {
                            if (!pendingMassRemoval) {
                              setPendingMassRemoval(true)
                              return
                            }

                            if (!selectedGroupKey) {
                              alert('Ingen faggruppe er valgt')
                              return
                            }

                            const [oldSubjectCode, oldGroupCode, oldBlock] = selectedGroupKey.split('-')
                            const oldSubjectName =
                              parsedData?.subjects.find((subject) => subject.code === oldSubjectCode)?.name || oldSubjectCode
                            const manualChanges: BalanceChange[] = []

                            const updatedStudents = parsedData!.students.map((student) => {
                              if (!selectedStudentsForMassUpdate.has(student.id)) {
                                return student
                              }

                              const oldAssignment = student.assignments.find(
                                (assignment) =>
                                  assignment.subjectCode === oldSubjectCode &&
                                  assignment.groupCode === oldGroupCode &&
                                  assignment.block === oldBlock,
                              )
                              if (!oldAssignment) {
                                return student
                              }

                              manualChanges.push({
                                studentId: student.id,
                                studentName: student.fullName,
                                subjectCode: oldSubjectCode,
                                subjectName: oldAssignment.subjectName,
                                fromGroupCode: oldAssignment.groupCode,
                                fromBlock: oldAssignment.block,
                                toGroupCode: '',
                                toBlock: '',
                              })

                              return {
                                ...student,
                                assignments: student.assignments.filter(
                                  (assignment) =>
                                    !(
                                      assignment.subjectCode === oldSubjectCode &&
                                      assignment.groupCode === oldGroupCode &&
                                      assignment.block === oldBlock
                                    ),
                                ),
                              }
                            })

                            const {
                              groupBreakdowns: newGroupBreakdowns,
                              blockBreakdowns: newBlockBreakdowns,
                            } = recalculateBreakdowns(updatedStudents, parsedData.groupBreakdowns)

                            const updatedData = {
                              ...parsedData!,
                              students: updatedStudents,
                              groupBreakdowns: newGroupBreakdowns,
                              blockBreakdowns: newBlockBreakdowns,
                            }
                            setParsedData(updatedData)

                            appendBalanceHistoryRun(
                              manualChanges,
                              `Masseoppdatering: ${manualChanges.length} elever fjernet fra ${oldSubjectName} (${oldGroupCode}) i ${oldBlock}`,
                            )

                            setBalanceMessage(
                              `Masseoppdatering: ${selectedStudentsForMassUpdate.size} elever fjernet fra ${oldSubjectName} (${oldGroupCode}) i ${oldBlock}`,
                            )

                            setSelectedStudentsForMassUpdate(new Set())
                            setShowMassUpdateDialog(false)
                            setPendingMassRemoval(false)
                            setMassUpdateTargetSubject('')
                            setMassUpdateTargetBlock('')
                            setSelectedGroupKey('')
                          }}
                          className="clear-results-button"
                        >
                          {pendingMassRemoval ? 'Bekreft fjern' : 'Fjern'}
                        </button>
                        <button
                          type="button"
                          onClick={() => {
                            setShowMassUpdateDialog(false)
                            setPendingMassRemoval(false)
                            setMassUpdateTargetSubject('')
                            setMassUpdateTargetBlock('')
                          }}
                          className="clear-results-button"
                        >
                          Avbryt
                        </button>
                      </div>
                    </div>
                  </div>
                </div>
              )}

              {showAddSubjectDialog && (
                <div
                  className="modal-overlay"
                  onClick={() => {
                    setShowAddSubjectDialog(false)
                    setAddSubjectTargetCode('')
                    setAddSubjectTargetBlock('')
                  }}
                >
                  <div className="modal-content" onClick={(e) => e.stopPropagation()}>
                    <h3>Legg til fag</h3>
                    <p>Legg til fag fra originalfilen i en blokk som ikke er brukt ennå.</p>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '1rem', marginTop: '1rem' }}>
                      <label>
                        <strong>Fag:</strong>
                        <select
                          value={addSubjectTargetCode}
                          onChange={(e) => {
                            setAddSubjectTargetCode(e.target.value)
                            setAddSubjectTargetBlock('')
                          }}
                          style={{ width: '100%', marginTop: '0.5rem', padding: '0.5rem' }}
                        >
                          <option value="">Velg fag...</option>
                          {parsedData.subjects
                            .filter((subject) => {
                              const originalBlocks = parsedData.originalAvailableBlocksBySubject[subject.code]
                              if (!originalBlocks) {
                                return false
                              }

                              const activeBlocks = new Set(
                                parsedData.groupBreakdowns
                                  .filter((group) => group.subjectCode === subject.code)
                                  .map((group) => group.block),
                              )

                              const hasUnusedBlock = originalBlocks.some((block) => !activeBlocks.has(block))
                              return hasUnusedBlock
                            })
                            .slice()
                            .sort((a, b) => a.name.localeCompare(b.name))
                            .map((subject) => (
                              <option key={subject.code} value={subject.code}>
                                {subject.code} - {subject.name}
                              </option>
                            ))}
                        </select>
                      </label>

                      <label>
                        <strong>Blokk:</strong>
                        <select
                          value={addSubjectTargetBlock}
                          onChange={(e) => setAddSubjectTargetBlock(e.target.value)}
                          style={{ width: '100%', marginTop: '0.5rem', padding: '0.5rem' }}
                        >
                          <option value="">Velg blokk...</option>
                          {addSubjectTargetCode && parsedData.originalAvailableBlocksBySubject[addSubjectTargetCode]
                            ? parsedData.originalAvailableBlocksBySubject[addSubjectTargetCode]
                                .filter((block) => {
                                  const hasGroupInBlock = parsedData.groupBreakdowns.some(
                                    (group) => group.subjectCode === addSubjectTargetCode && group.block === block,
                                  )
                                  return !hasGroupInBlock
                                })
                                .map((block) => (
                                  <option key={block} value={block}>
                                    {block}
                                  </option>
                                ))
                            : null}
                        </select>
                      </label>

                      <div style={{ display: 'flex', gap: '0.5rem', marginTop: '0.5rem' }}>
                        <button
                          type="button"
                          onClick={() => {
                            if (!addSubjectTargetCode || !addSubjectTargetBlock) {
                              alert('Vennligst velg både fag og blokk')
                              return
                            }

                            const targetSubject = parsedData.subjects.find((subject) => subject.code === addSubjectTargetCode)
                            if (!targetSubject) {
                              alert('Fag ikke funnet')
                              return
                            }

                            const originalBlocks = parsedData.originalAvailableBlocksBySubject[addSubjectTargetCode]
                            if (!originalBlocks || !originalBlocks.includes(addSubjectTargetBlock)) {
                              alert('Valgt blokk er ikke tilgjengelig for dette faget i originalfilen')
                              return
                            }

                            const hasGroupInTargetBlock = parsedData.groupBreakdowns.some(
                              (group) => group.subjectCode === addSubjectTargetCode && group.block === addSubjectTargetBlock,
                            )

                            const wasAlreadyAvailable = hasGroupInTargetBlock
                            if (wasAlreadyAvailable) {
                              alert(`${targetSubject.name} er allerede tilgjengelig i ${addSubjectTargetBlock}.`)
                              setShowAddSubjectDialog(false)
                              setAddSubjectTargetCode('')
                              setAddSubjectTargetBlock('')
                              return
                            }

                            const updatedSubjects = parsedData.subjects.map((subject) => {
                              if (subject.code !== addSubjectTargetCode) {
                                return subject
                              }

                              const nextBlocks = sortBlocks([...subject.blocks, addSubjectTargetBlock])

                              return {
                                ...subject,
                                blocks: nextBlocks,
                              }
                            })

                            const updatedBlocks = parsedData.blocks.includes(addSubjectTargetBlock)
                              ? parsedData.blocks
                              : sortBlocks([...parsedData.blocks, addSubjectTargetBlock])

                            const updatedGroupBreakdowns = hasGroupInTargetBlock
                              ? parsedData.groupBreakdowns
                              : parsedData.groupBreakdowns.concat([
                                  {
                                    subjectCode: addSubjectTargetCode,
                                    subjectName: targetSubject.name,
                                    groupCode: createDefaultGroupCode(addSubjectTargetCode, addSubjectTargetBlock),
                                    block: addSubjectTargetBlock,
                                    studentCount: 0,
                                  },
                                ])

                            const updatedData = {
                              ...parsedData,
                              subjects: updatedSubjects,
                              blocks: updatedBlocks,
                              groupBreakdowns: updatedGroupBreakdowns,
                            }
                            setParsedData(updatedData)

                            // Log to balance history like a balance operation
                            const addSubjectChange: BalanceChange = {
                              studentId: 'SYSTEM',
                              studentName: 'System',
                              subjectCode: addSubjectTargetCode,
                              subjectName: targetSubject.name,
                              fromGroupCode: '',
                              fromBlock: '',
                              toGroupCode: createDefaultGroupCode(addSubjectTargetCode, addSubjectTargetBlock),
                              toBlock: addSubjectTargetBlock,
                            }

                            const newHistoryRun: BalanceResultRun = {
                              id: crypto.randomUUID(),
                              createdAt: new Date().toISOString(),
                              message: `Lagt til ${targetSubject.name} i ${addSubjectTargetBlock}`,
                              changes: [addSubjectChange],
                            }

                            const updatedBalanceHistory = [...balanceHistory, newHistoryRun]
                            setBalanceHistory(updatedBalanceHistory)
                            setBalanceMessage(`${targetSubject.name} er nå tilgjengelig i ${addSubjectTargetBlock}.`)

                            setShowAddSubjectDialog(false)
                            setAddSubjectTargetCode('')
                            setAddSubjectTargetBlock('')
                          }}
                          className="balance-button"
                          disabled={!addSubjectTargetCode || !addSubjectTargetBlock}
                        >
                          Legg til
                        </button>
                        <button
                          type="button"
                          onClick={() => {
                            setShowAddSubjectDialog(false)
                            setAddSubjectTargetCode('')
                            setAddSubjectTargetBlock('')
                          }}
                          className="clear-results-button"
                        >
                          Avbryt
                        </button>
                      </div>
                    </div>
                  </div>
                </div>
              )}

              {showProgressiveBalanceDialog && (
                <div className="modal-overlay" onClick={() => setShowProgressiveBalanceDialog(false)}>
                  <div className="modal-content" onClick={(e) => e.stopPropagation()}>
                    <h3>Progressiv balansering</h3>
                    <p>Velg maksimalt antall elever under kapasitet å tillate før prøving av neste offset.</p>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '1rem', marginTop: '1rem' }}>
                      <div>
                        <label>
                          <strong>Maksimalt offset under kapasitet:</strong>
                        </label>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '1rem', marginTop: '0.5rem' }}>
                          <input
                            type="range"
                            min={-15}
                            max={0}
                            value={progressiveBalanceMaxOffset}
                            onChange={(e) => setProgressiveBalanceMaxOffset(parseInt(e.target.value, 10))}
                            style={{ flex: 1 }}
                          />
                          <span style={{ minWidth: '3rem', textAlign: 'right' }}>
                            <strong>{progressiveBalanceMaxOffset}</strong>
                          </span>
                        </div>
                        <p style={{ fontSize: '0.85rem', color: '#666', marginTop: '0.5rem' }}>
                          {progressiveBalanceMaxOffset === 0
                            ? 'Balansering til full kapasitet'
                            : `Tillater inntil ${Math.abs(progressiveBalanceMaxOffset)} færre elever enn kapasitet før neste versøk`}
                        </p>
                      </div>

                      <div style={{ display: 'flex', gap: '0.5rem', marginTop: '0.5rem' }}>
                        <button
                          type="button"
                          onClick={() => {
                            if (!parsedData) return

                            const result = progressiveBalanceGroups(parsedData.students, parsedData.groupBreakdowns, progressiveBalanceMaxOffset, (groups) => setDebugGroups(groups))
                            setBalanceResults(result.allChanges)
                            
                            // Apply changes to the parsed data
                            if (result.allChanges.length > 0) {
                              const updatedStudents = applyBalanceChanges(parsedData.students, result.allChanges)
                              const { groupBreakdowns: newGroupBreakdowns, blockBreakdowns: newBlockBreakdowns } = recalculateBreakdowns(updatedStudents, parsedData.groupBreakdowns)
                              const updatedData = {
                                ...parsedData,
                                students: updatedStudents,
                                groupBreakdowns: newGroupBreakdowns,
                                blockBreakdowns: newBlockBreakdowns,
                              }
                              setParsedData(updatedData)
                            }
                            
                            setBalanceMessage(result.summary)
                            appendBalanceHistoryRun(result.allChanges, result.summary, result.collisionErrors)
                            setShowProgressiveBalanceDialog(false)
                          }}
                          className="balance-button"
                        >
                          Kjør progressiv balansering
                        </button>
                        <button
                          type="button"
                          onClick={() => setShowProgressiveBalanceDialog(false)}
                          className="clear-results-button"
                        >
                          Avbryt
                        </button>
                      </div>
                    </div>
                  </div>
                </div>
              )}



            </section>
          ) : viewMode === 'blokkoversikt' ? (
            <section className="subject-panel">
              <div className="subject-view-header">
                <h2>Blokkoversikt</h2>
              </div>

              <table>
                <thead>
                  <tr>
                    <th>Fag</th>
                    <th>Tittel</th>
                    <th>Blokker</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredSubjects
                    .slice()
                    .sort((a, b) => a.name.localeCompare(b.name))
                    .map((subject) => {
                      const sortedBlocks = sortBlocks(
                        Array.from(blokkoversiktBlocksBySubject.get(subject.code) || new Set<string>()),
                      )

                      if (blockFilter && !sortedBlocks.includes(blockFilter)) {
                        return null
                      }

                      if (onlyBlokkfag && sortedBlocks.length === 0) {
                        return null
                      }

                      return (
                        <tr key={`${subject.code}-blokkoversikt`}>
                          <td>{subject.code}</td>
                          <td>{subject.name}</td>
                          <td>{sortedBlocks.length > 0 ? sortedBlocks.join(', ') : '-'}</td>
                        </tr>
                      )
                    })}
                </tbody>
              </table>
            </section>
          ) : (
            <section className="subject-panel">
              <div className="subject-view-header">
                <h2>Blokkvalg</h2>
              </div>

              {(() => {
                const getSubjectColor = (subjectName: string, useLighterShade: boolean): string => {
                  const name = subjectName.toLowerCase()
                  // Science subjects - Green
                  if (name.includes('biologi') || name.includes('fysikk') || name.includes('kjemi') || 
                      name.includes('geofag') || name.includes('matematikk')) {
                    return useLighterShade ? '#d9f0d9' : '#a8d5a8'
                  }
                  // Sports - Orange
                  if (name.includes('idrett') || name.includes('friluftsliv')) {
                    return useLighterShade ? '#ffe4c7' : '#ffc896'
                  }
                  // Social sciences/languages - Blue
                  return useLighterShade ? '#dcecff' : '#b3d9ff'
                }

                const getSubjectSortRank = (subjectName: string): number => {
                  const name = subjectName.toLowerCase()
                  const isScience =
                    name.includes('biologi') ||
                    name.includes('fysikk') ||
                    name.includes('kjemi') ||
                    name.includes('geofag') ||
                    name.includes('matematikk')
                  const isSports = name.includes('idrett') || name.includes('friluftsliv')

                  if (isScience) {
                    return 0
                  }
                  if (isSports) {
                    return 2
                  }
                  // Social sciences and languages are in the middle.
                  return 1
                }

                const getDisplaySubjectName = (subjectName: string): string => {
                  const normalized = normalizeSubjectName(subjectName)
                  const lower = subjectName.toLowerCase()

                  if (lower.includes('entreprenørskap og bedriftsutvikling') || lower.includes('entreprenorskap og bedriftsutvikling')) {
                    const levelMatch = subjectName.match(/([12])\s*$/u)
                    return levelMatch ? `Entreprenørskap ${levelMatch[1]}` : 'Entreprenørskap'
                  }
                  if (lower.includes('markedsføring og ledelse') || lower.includes('markedsforing og ledelse')) {
                    const levelMatch = subjectName.match(/([12])\s*$/u)
                    return levelMatch ? `Markedsføring ${levelMatch[1]}` : 'Markedsføring'
                  }
                  if (normalized === 'internasjonal engelsk, skriftlig') {
                    return 'Engelsk 1'
                  }
                  if (normalized === 'samfunnsfaglig engelsk, skriftlig') {
                    return 'Engelsk 2'
                  }
                  if (normalized === 'sosiologi og sosialantropologi') {
                    return 'Sosiologi'
                  }

                  return subjectName
                }

                const formatSubjectLabel = (subjectName: string): string => {
                  const normalized = subjectName.trim()
                  if (normalized.length <= 10) {
                    return normalized
                  }

                  const trailingLevel = normalized.match(/([12])\s*$/u)
                  if (trailingLevel) {
                    const suffix = ` ${trailingLevel[1]}`
                    const prefixLength = Math.max(1, 10 - suffix.length)
                    const prefix = normalized.slice(0, prefixLength).trimEnd()
                    return `${prefix}${suffix}`
                  }

                  return normalized.slice(0, 10)
                }

                const normalizeSubjectName = (value: string): string =>
                  value
                    .toLowerCase()
                    .normalize('NFD')
                    .replace(/[\u0300-\u036f]/g, '')
                    .replace(/\s+/g, ' ')
                    .trim()

                const vg3OnlySubjects = new Set<string>([
                  'biologi 2',
                  'fransk niva iii',
                  'fysikk 2',
                  'geofag 2',
                  'kjemi 2',
                  'markedsforing og ledelse 2',
                  'matematikk r2',
                  'matematikk s2',
                  'politikk og menneskerettigheter',
                  'psykologi 2',
                  'rettslaere 2',
                  'samfunnsfaglig engelsk',
                  'samfunnsokonomi 2',
                  'spansk niva iii',
                  'toppidrett 3',
                  'tysk niva iii',
                ])

                const isVg3OnlySubject = (subjectName: string): boolean =>
                  vg3OnlySubjects.has(normalizeSubjectName(subjectName))

                // Only show blokkfag in Bytteoversikt and its options.
                const allSubjects = (parsedData?.subjects || [])
                  .filter((subject) => Array.isArray(subject.blocks) && subject.blocks.length > 0)
                  .filter((subject) => subject.blocks.some((block) => !block.toLowerCase().includes('matte')))
                  .slice()
                  .sort((a, b) => a.name.localeCompare(b.name))

                // Get all unique blocks from all subjects, sorted numerically
                const allBlocks = Array.from(
                  new Set(allSubjects.flatMap(s => s.blocks || []))
                )
                .filter(block => !block.toLowerCase().includes('matte')) // Hide matte block
                .sort((a, b) => {
                  const numA = parseInt(a)
                  const numB = parseInt(b)
                  if (!isNaN(numA) && !isNaN(numB)) {
                    return numA - numB
                  }
                  return a.localeCompare(b)
                })
                
                const renderBlockRows = (section: 'vg2' | 'vg3') => (
                  <>
                    {allBlocks
                      .filter((block) => {
                        const normalized = block.trim().toLowerCase()
                        const blockNumberMatch = normalized.match(/(?:blokk\s*)?(\d+)/u)
                        const blockNumber = blockNumberMatch ? blockNumberMatch[1] : normalized

                        if (section === 'vg2' && blockNumber === '4') {
                          return false
                        }
                        if (section === 'vg3' && blockNumber === '1') {
                          return false
                        }
                        return true
                      })
                      .map((block) => {
                      // Find all subjects that have this block
                      const subjectsInBlock = allSubjects
                        .filter((subject) => subject.blocks && subject.blocks.includes(block))
                        .filter((subject) => (section === 'vg2' ? !isVg3OnlySubject(subject.name) : true))
                        .filter((subject) => (bytteSubjectVisibility[subject.code]?.[section] ?? true))
                        .sort((a, b) => {
                          const rankDiff = getSubjectSortRank(a.name) - getSubjectSortRank(b.name)
                          if (rankDiff !== 0) {
                            return rankDiff
                          }
                          
                          // Sort by VG2 visibility (lighter colors first)
                          const aVisibility = bytteSubjectVisibility[a.code] ?? { vg2: true, vg3: true }
                          const bVisibility = bytteSubjectVisibility[b.code] ?? { vg2: true, vg3: true }
                          const aIsVg3Only = isVg3OnlySubject(a.name)
                          const bIsVg3Only = isVg3OnlySubject(b.name)
                          
                          const aIsLighter = !aIsVg3Only && aVisibility.vg2
                          const bIsLighter = !bIsVg3Only && bVisibility.vg2
                          
                          if (aIsLighter !== bIsLighter) {
                            return aIsLighter ? -1 : 1  // Lighter colors first
                          }
                          
                          return a.name.localeCompare(b.name)
                        })

                      return (
                        <div key={block} style={{
                          display: 'grid',
                          gridTemplateColumns: '100px 1fr',
                          gap: '8px',
                          marginBottom: '8px',
                          alignItems: 'start'
                        }}>
                          <div style={{
                            backgroundColor: '#b0b0b0',
                            padding: '6px',
                            fontWeight: 'bold',
                            textAlign: 'center',
                            minHeight: '50px',
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: 'center',
                            fontSize: '0.85rem'
                          }}>
                            {block.toLowerCase().startsWith('blokk') ? block : `Blokk ${block}`}
                          </div>
                          <div style={{
                            display: 'flex',
                            flexWrap: 'wrap',
                            gap: '6px',
                            alignContent: 'start'
                          }}>
                            {subjectsInBlock.map((subject) => {
                              const visibility = bytteSubjectVisibility[subject.code] ?? { vg2: true, vg3: true }
                              const isVg3Only = isVg3OnlySubject(subject.name)
                              const useLighterShade = !isVg3Only && visibility.vg2
                              
                              return (
                                <div
                                  key={subject.code}
                                  style={{
                                    border: '2px solid #333',
                                    padding: '6px 10px',
                                    backgroundColor: getSubjectColor(subject.name, useLighterShade),
                                    minWidth: '80px',
                                    minHeight: '40px',
                                    display: 'flex',
                                    flexDirection: 'column',
                                    justifyContent: 'center'
                                  }}
                                >
                                  <div style={{ fontWeight: 'bold', fontSize: '0.75rem', lineHeight: 1.2 }}>
                                    {getDisplaySubjectName(subject.name)}
                                  </div>
                                </div>
                              )
                            })}
                          </div>
                        </div>
                      )
                    })}
                  </>
                )

                return (
                  <div>
                    <div style={{ marginBottom: '0.75rem', border: '1px solid #d0d7de', borderRadius: '8px', padding: '0.5rem' }}>
                      <h3 
                        style={{ 
                          marginTop: 0, 
                          marginBottom: bytteOptionsExpanded ? '0.5rem' : 0, 
                          cursor: 'pointer',
                          userSelect: 'none',
                          display: 'flex',
                          alignItems: 'center',
                          gap: '0.5rem'
                        }}
                        onClick={() => setBytteOptionsExpanded(!bytteOptionsExpanded)}
                      >
                        <span style={{ fontSize: '0.9rem' }}>{bytteOptionsExpanded ? '▼' : '▶'}</span>
                        Options
                      </h3>
                      {bytteOptionsExpanded && (
                        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(128px, 1fr))', gap: '0.4rem', maxHeight: '220px', overflowY: 'auto' }}>
                          {allSubjects.map((subject) => {
                            const visibility = bytteSubjectVisibility[subject.code] ?? { vg2: true, vg3: true }
                            const isVg3Only = isVg3OnlySubject(subject.name)
                            return (
                              <div key={`${subject.code}-row`} style={{ border: '1px solid #d0d7de', borderRadius: '6px', padding: '0.35rem 0.4rem', backgroundColor: '#fff' }}>
                                <div style={{ fontSize: '0.78rem', fontWeight: 600, marginBottom: '0.3rem', lineHeight: 1.2 }}>
                                  {formatSubjectLabel(getDisplaySubjectName(subject.name))}
                                </div>
                                <div style={{ display: 'grid', gridTemplateColumns: isVg3Only ? '1fr' : '1fr 1fr', gap: '0.25rem' }}>
                                  {!isVg3Only && (
                                    <label style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '0.2rem', fontSize: '0.72rem' }}>
                                      <input
                                        type="checkbox"
                                        checked={visibility.vg2}
                                        onChange={(event) => {
                                          const checked = event.target.checked
                                          setBytteSubjectVisibility((previous) => ({
                                            ...previous,
                                            [subject.code]: {
                                              ...(previous[subject.code] ?? { vg2: true, vg3: true }),
                                              vg2: checked,
                                            },
                                          }))
                                        }}
                                      />
                                      VG2
                                    </label>
                                  )}
                                  <label style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '0.2rem', fontSize: '0.72rem' }}>
                                    <input
                                      type="checkbox"
                                      checked={visibility.vg3}
                                      onChange={(event) => {
                                        const checked = event.target.checked
                                        setBytteSubjectVisibility((previous) => ({
                                          ...previous,
                                          [subject.code]: {
                                            ...(previous[subject.code] ?? { vg2: true, vg3: true }),
                                            vg3: checked,
                                          },
                                        }))
                                      }}
                                    />
                                    VG3
                                  </label>
                                </div>
                              </div>
                            )
                          })}
                        </div>
                      )}
                    </div>
                    <h3 style={{ marginBottom: '0.75rem' }}>VG2</h3>
                    {renderBlockRows('vg2')}
                    <h3 style={{ marginTop: '1.5rem', marginBottom: '0.75rem' }}>VG3</h3>
                    {renderBlockRows('vg3')}
                  </div>
                )
              })()}
            </section>
          )}

          {showStudentAddSubjectDialog && selectedStudent && (
            <div
              className="modal-overlay"
              onClick={() => {
                setShowStudentAddSubjectDialog(false)
                setStudentAddSubjectCode('')
                setStudentAddSubjectBlock('')
              }}
            >
              <div className="modal-content" onClick={(e) => e.stopPropagation()}>
                <h3>Legg til fag for {selectedStudent.fullName}</h3>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '1rem', marginTop: '1rem' }}>
                  <label>
                    <strong>Velg fag:</strong>
                    <select
                      value={studentAddSubjectCode}
                      onChange={(e) => {
                        setStudentAddSubjectCode(e.target.value)
                        setStudentAddSubjectBlock('')
                      }}
                      style={{ width: '100%', marginTop: '0.5rem', padding: '0.5rem' }}
                    >
                      <option value="">Velg fag...</option>
                      {parsedData.subjects
                        .filter((subject) => subject.blocks.length > 0)
                        .slice()
                        .sort((a, b) => a.name.localeCompare(b.name))
                        .map((subject) => (
                          <option key={subject.code} value={subject.code}>
                            {subject.code} - {subject.name}
                          </option>
                        ))}
                    </select>
                  </label>

                  {studentAddSubjectCode && (
                    <div>
                      <strong>Velg blokk:</strong>
                      <select
                        value={studentAddSubjectBlock}
                        onChange={(e) => setStudentAddSubjectBlock(e.target.value)}
                        style={{ width: '100%', marginTop: '0.5rem', padding: '0.5rem' }}
                      >
                        <option value="">Velg blokk...</option>
                        {getSubjectBlockChoices(parsedData, studentAddSubjectCode).map((choice) => (
                          <option key={choice.block} value={choice.block}>
                            {choice.block} ({choice.studentCount} elever)
                          </option>
                        ))}
                      </select>
                    </div>
                  )}

                  <div style={{ display: 'flex', gap: '0.5rem', marginTop: '0.5rem' }}>
                    <button
                      type="button"
                      onClick={() => {
                        if (!studentAddSubjectCode || !studentAddSubjectBlock) {
                          alert('Vennligst velg fag og blokk')
                          return
                        }

                        const block = studentAddSubjectBlock
                        const targetSubject = parsedData.subjects.find((s) => s.code === studentAddSubjectCode)
                        if (!targetSubject) {
                          alert('Fag ikke funnet')
                          return
                        }

                        const targetGroups = parsedData.groupBreakdowns
                          .filter((group) => group.subjectCode === studentAddSubjectCode && group.block === block)
                          .sort((a, b) => a.studentCount - b.studentCount)

                        const smallestGroup = targetGroups[0]
                        if (!smallestGroup) {
                          alert(`Ingen grupper funnet for ${targetSubject.name} i ${block}`)
                          return
                        }

                        const groupCode = smallestGroup.groupCode

                        // Check if student already has this subject
                        if (selectedStudent.assignments.some((a) => a.subjectCode === studentAddSubjectCode)) {
                          alert(`${selectedStudent.fullName} har allerede ${targetSubject.name}`)
                          return
                        }

                        // Add the subject to the student
                        const updatedStudents = parsedData.students.map((student) => {
                          if (student.id !== selectedStudent.id) {
                            return student
                          }

                          return {
                            ...student,
                            assignments: [
                              ...student.assignments,
                              {
                                subjectCode: studentAddSubjectCode,
                                subjectName: targetSubject.name,
                                groupCode,
                                block,
                              },
                            ],
                          }
                        })

                        // Recalculate breakdowns
                        const { groupBreakdowns, blockBreakdowns } = recalculateBreakdowns(updatedStudents, parsedData.groupBreakdowns)

                        // Update subjects with new student count
                        const updatedSubjects = parsedData.subjects.map((subject) => {
                          if (subject.code !== studentAddSubjectCode) {
                            return subject
                          }
                          return {
                            ...subject,
                            studentCount: subject.studentCount + 1,
                          }
                        })

                        setParsedData({
                          ...parsedData,
                          students: updatedStudents,
                          subjects: updatedSubjects,
                          groupBreakdowns,
                          blockBreakdowns,
                        })

                        appendBalanceHistoryRun(
                          [
                            {
                              studentId: selectedStudent.id,
                              studentName: selectedStudent.fullName,
                              subjectCode: studentAddSubjectCode,
                              subjectName: targetSubject.name,
                              fromGroupCode: '',
                              fromBlock: '',
                              toGroupCode: groupCode,
                              toBlock: block,
                            },
                          ],
                          `${targetSubject.name} (${groupCode}) lagt til for ${selectedStudent.fullName}`,
                        )

                        setBalanceMessage(`${targetSubject.name} (${groupCode}) lagt til for ${selectedStudent.fullName}`)
                        setShowStudentAddSubjectDialog(false)
                        setStudentAddSubjectCode('')
                        setStudentAddSubjectBlock('')
                      }}
                      className="balance-button"
                      disabled={!studentAddSubjectCode || !studentAddSubjectBlock}
                    >
                      Legg til
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        setShowStudentAddSubjectDialog(false)
                        setStudentAddSubjectCode('')
                        setStudentAddSubjectBlock('')
                      }}
                      className="clear-results-button"
                    >
                      Avbryt
                    </button>
                  </div>
                </div>
              </div>
            </div>
          )}

          {showStudentSwapDialog && selectedStudent && parsedData && (
            <div
              className="modal-overlay"
              onClick={() => {
                setShowStudentSwapDialog(false)
                setStudentSwapAssignmentKey('')
                setStudentSwapTargetSubject('')
                setStudentSwapTargetBlock('')
              }}
            >
              <div className="modal-content" onClick={(e) => e.stopPropagation()}>
                {(() => {
                  const [subjectCode, groupCode, block] = studentSwapAssignmentKey.split('|')
                  const swapAssignment = selectedStudent.assignments.find(
                    (assignment) =>
                      assignment.subjectCode === subjectCode
                      && assignment.groupCode === groupCode
                      && assignment.block === block,
                  )

                  if (!swapAssignment) {
                    return (
                      <>
                        <h3>Bytt blokk</h3>
                        <p>Fant ikke valgt fag for eleven. Lukk og prøv igjen.</p>
                      </>
                    )
                  }

                  const targetSubjectCode = studentSwapTargetSubject || swapAssignment.subjectCode
                  const targetSubject = parsedData.subjects.find((subject) => subject.code === targetSubjectCode)
                  const blockChoices = getSubjectBlockChoices(parsedData, targetSubjectCode)

                  return (
                    <>
                      <h3>Bytt fag/blokk for {selectedStudent.fullName}</h3>
                      <p style={{ marginTop: '0.4rem' }}>
                        <strong>{swapAssignment.subjectName}</strong> er nå i {swapAssignment.block} ({swapAssignment.groupCode})
                      </p>

                      <div style={{ display: 'flex', flexDirection: 'column', gap: '1rem', marginTop: '1rem' }}>
                        <label>
                          <strong>Nytt fag:</strong>
                          <select
                            value={targetSubjectCode}
                            onChange={(e) => {
                              setStudentSwapTargetSubject(e.target.value)
                              setStudentSwapTargetBlock('')
                            }}
                            style={{ width: '100%', marginTop: '0.5rem', padding: '0.5rem' }}
                          >
                            <option value="">Velg fag...</option>
                            {parsedData.subjects
                              .filter((subject) => subject.blocks.length > 0)
                              .slice()
                              .sort((a, b) => a.name.localeCompare(b.name))
                              .map((subject) => (
                                <option key={subject.code} value={subject.code}>
                                  {subject.code} - {subject.name}
                                </option>
                              ))}
                          </select>
                        </label>

                        <label>
                          <strong>Ny blokk:</strong>
                          <select
                            value={studentSwapTargetBlock}
                            onChange={(e) => setStudentSwapTargetBlock(e.target.value)}
                            style={{ width: '100%', marginTop: '0.5rem', padding: '0.5rem' }}
                          >
                            <option value="">Velg blokk...</option>
                            {blockChoices.map((choice) => (
                              <option key={choice.block} value={choice.block}>
                                {choice.block} ({choice.studentCount} elever)
                              </option>
                            ))}
                          </select>
                        </label>

                        <div style={{ display: 'flex', gap: '0.5rem' }}>
                          <button
                            type="button"
                            onClick={() => {
                              if (!targetSubjectCode) {
                                alert('Vennligst velg fag')
                                return
                              }

                              if (!studentSwapTargetBlock) {
                                alert('Vennligst velg ny blokk')
                                return
                              }

                              if (!targetSubject) {
                                alert('Fag ikke funnet')
                                return
                              }

                              if (
                                targetSubjectCode !== swapAssignment.subjectCode
                                && selectedStudent.assignments.some((assignment) => assignment.subjectCode === targetSubjectCode)
                              ) {
                                alert(`${selectedStudent.fullName} har allerede ${targetSubject.name}`)
                                return
                              }

                              const targetGroups = parsedData.groupBreakdowns
                                .filter((group) => group.subjectCode === targetSubjectCode && group.block === studentSwapTargetBlock)
                                .sort((a, b) => a.studentCount - b.studentCount)

                              const smallestGroup = targetGroups[0]
                              if (!smallestGroup) {
                                alert(`Ingen grupper funnet for ${targetSubject.name} i ${studentSwapTargetBlock}`)
                                return
                              }

                              const updatedStudents = parsedData.students.map((student) => {
                                if (student.id !== selectedStudent.id) {
                                  return student
                                }

                                if (targetSubjectCode !== swapAssignment.subjectCode) {
                                  return {
                                    ...student,
                                    assignments: student.assignments
                                      .filter((assignment) => !(
                                        assignment.subjectCode === swapAssignment.subjectCode
                                        && assignment.groupCode === swapAssignment.groupCode
                                        && assignment.block === swapAssignment.block
                                      ))
                                      .concat([
                                        {
                                          subjectCode: targetSubjectCode,
                                          subjectName: targetSubject.name,
                                          groupCode: smallestGroup.groupCode,
                                          block: studentSwapTargetBlock,
                                        },
                                      ]),
                                  }
                                }

                                return {
                                  ...student,
                                  assignments: student.assignments.map((assignment) => {
                                    if (
                                      assignment.subjectCode === swapAssignment.subjectCode
                                      && assignment.groupCode === swapAssignment.groupCode
                                      && assignment.block === swapAssignment.block
                                    ) {
                                      return {
                                        ...assignment,
                                        subjectName: targetSubject.name,
                                        groupCode: smallestGroup.groupCode,
                                        block: studentSwapTargetBlock,
                                      }
                                    }
                                    return assignment
                                  }),
                                }
                              })

                              const { groupBreakdowns, blockBreakdowns } = recalculateBreakdowns(updatedStudents, parsedData.groupBreakdowns)

                              const updatedSubjects = parsedData.subjects.map((subject) => {
                                if (targetSubjectCode === swapAssignment.subjectCode) {
                                  return subject
                                }

                                if (subject.code === swapAssignment.subjectCode) {
                                  return {
                                    ...subject,
                                    studentCount: Math.max(0, subject.studentCount - 1),
                                  }
                                }

                                if (subject.code === targetSubjectCode) {
                                  return {
                                    ...subject,
                                    studentCount: subject.studentCount + 1,
                                  }
                                }

                                return subject
                              })

                              setParsedData({
                                ...parsedData,
                                students: updatedStudents,
                                subjects: updatedSubjects,
                                groupBreakdowns,
                                blockBreakdowns,
                              })

                              const manualChanges: BalanceChange[] = targetSubjectCode === swapAssignment.subjectCode
                                ? [
                                    {
                                      studentId: selectedStudent.id,
                                      studentName: selectedStudent.fullName,
                                      subjectCode: swapAssignment.subjectCode,
                                      subjectName: targetSubject.name,
                                      fromGroupCode: swapAssignment.groupCode,
                                      fromBlock: swapAssignment.block,
                                      toGroupCode: smallestGroup.groupCode,
                                      toBlock: studentSwapTargetBlock,
                                    },
                                  ]
                                : [
                                    {
                                      studentId: selectedStudent.id,
                                      studentName: selectedStudent.fullName,
                                      subjectCode: swapAssignment.subjectCode,
                                      subjectName: swapAssignment.subjectName,
                                      fromGroupCode: swapAssignment.groupCode,
                                      fromBlock: swapAssignment.block,
                                      toGroupCode: '',
                                      toBlock: '',
                                    },
                                    {
                                      studentId: selectedStudent.id,
                                      studentName: selectedStudent.fullName,
                                      subjectCode: targetSubjectCode,
                                      subjectName: targetSubject.name,
                                      fromGroupCode: '',
                                      fromBlock: '',
                                      toGroupCode: smallestGroup.groupCode,
                                      toBlock: studentSwapTargetBlock,
                                    },
                                  ]

                              const nextMessage = targetSubjectCode === swapAssignment.subjectCode
                                ? `${swapAssignment.subjectName} byttet for ${selectedStudent.fullName}: ${swapAssignment.block} -> ${studentSwapTargetBlock}`
                                : `${swapAssignment.subjectName} byttet til ${targetSubject.name} for ${selectedStudent.fullName}`

                              appendBalanceHistoryRun(manualChanges, nextMessage)

                              setBalanceMessage(nextMessage)
                              setShowStudentSwapDialog(false)
                              setStudentSwapAssignmentKey('')
                              setStudentSwapTargetSubject('')
                              setStudentSwapTargetBlock('')
                            }}
                            className="balance-button"
                            disabled={!targetSubjectCode || !studentSwapTargetBlock}
                          >
                            Bytt
                          </button>
                          <button
                            type="button"
                            onClick={() => {
                              setShowStudentSwapDialog(false)
                              setStudentSwapAssignmentKey('')
                              setStudentSwapTargetSubject('')
                              setStudentSwapTargetBlock('')
                            }}
                            className="clear-results-button"
                          >
                            Avbryt
                          </button>
                        </div>
                      </div>
                    </>
                  )
                })()}
              </div>
            </div>
          )}
        </>
      )}
    </div>
  )
}

export default App
