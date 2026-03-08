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

type BalanceResultRun = {
  id: string
  createdAt: string
  message: string
  changes: BalanceChange[]
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

type PersistedUiState = {
  selectedStudentId: string
  studentQuery: string
  subjectQuery: string
  blockFilter: string
  viewMode: 'students' | 'subjects'
  onlyBlokkfag: boolean
  showIncompleteBlocks: boolean
  showOverloadedStudents: boolean
  showBlockCollisions: boolean
  showDuplicateSubjects: boolean
}

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

  return { changes, overcrowdedCount: overcrowded.length, partnerLookaheadMoves }
}

function progressiveBalanceGroups(
  students: StudentRecord[],
  maxOffset: number,
  debugCallback?: (groups: Array<{ key: string; count: number; maxCap: number; status: string }>) => void
): { allChanges: BalanceChange[]; summary: string } {
  const allChanges: BalanceChange[] = []
  const offsets = []
  let totalPartnerLookaheadMoves = 0
  
  // Generate offsets from maxOffset down to 0
  for (let i = maxOffset; i <= 0; i++) {
    offsets.push(i)
  }

  // Apply balancing iteratively
  let currentStudents = students
  
  for (const offset of offsets) {
    const result = balanceGroupsWithOffset(currentStudents, offset, debugCallback)
    totalPartnerLookaheadMoves += result.partnerLookaheadMoves
    
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

  return { allChanges, summary }
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

  return { changes, overcrowdedCount: overcrowded.length }
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

function recalculateBreakdowns(students: StudentRecord[]): { groupBreakdowns: GroupBreakdownRecord[]; blockBreakdowns: BlockBreakdownRecord[] } {
  const groupSummary = new Map<string, { subjectCode: string; subjectName: string; groupCode: string; block: string; students: Set<string> }>()

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
    const assignment: Assignment = {
      subjectCode,
      subjectName: subjectsByCode.get(subjectCode) || subjectCode,
      groupCode,
      block: row.Blockname || inferredBlock,
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
  const [viewMode, setViewMode] = useState<'students' | 'subjects'>(persistedUiState.viewMode)
  const [selectedGroupKey, setSelectedGroupKey] = useState<string>('')
  const [onlyBlokkfag, setOnlyBlokkfag] = useState<boolean>(persistedUiState.onlyBlokkfag)
  const [showIncompleteBlocks, setShowIncompleteBlocks] = useState<boolean>(persistedUiState.showIncompleteBlocks)
  const [showOverloadedStudents, setShowOverloadedStudents] = useState<boolean>(persistedUiState.showOverloadedStudents)
  const [showBlockCollisions, setShowBlockCollisions] = useState<boolean>(persistedUiState.showBlockCollisions)
  const [showDuplicateSubjects, setShowDuplicateSubjects] = useState<boolean>(persistedUiState.showDuplicateSubjects)
  const [balanceResults, setBalanceResults] = useState<BalanceChange[] | null>(null)
  const [balanceHistory, setBalanceHistory] = useState<BalanceResultRun[]>(() => {
    const stored = loadFromLocalStorage<BalanceResultRun[]>(STORAGE_KEYS.balanceHistory, [])
    return Array.isArray(stored)
      ? stored.filter((run) => run && Array.isArray(run.changes) && typeof run.createdAt === 'string')
      : []
  })
  const [balanceMessage, setBalanceMessage] = useState<string>('')
  const [debugGroups, setDebugGroups] = useState<Array<{ key: string; count: number; maxCap: number; status: string }>>([])
  const [balanceDeltaCounts, setBalanceDeltaCounts] = useState<Map<string, number>>(new Map())
  const [balanceBlockDeltaCounts, setBalanceBlockDeltaCounts] = useState<Map<string, number>>(new Map())
  const [errorMessage, setErrorMessage] = useState<string>('')
  const [selectedStudentsForMassUpdate, setSelectedStudentsForMassUpdate] = useState<Set<string>>(new Set())
  const [showMassUpdateDialog, setShowMassUpdateDialog] = useState<boolean>(false)
  const [massUpdateTargetSubject, setMassUpdateTargetSubject] = useState<string>('')
  const [massUpdateTargetBlock, setMassUpdateTargetBlock] = useState<string>('')
  const [showAddSubjectDialog, setShowAddSubjectDialog] = useState<boolean>(false)
  const [addSubjectTargetCode, setAddSubjectTargetCode] = useState<string>('')
  const [addSubjectTargetBlock, setAddSubjectTargetBlock] = useState<string>('')
  const [showStudentAddSubjectDialog, setShowStudentAddSubjectDialog] = useState<boolean>(false)
  const [studentAddSubjectCode, setStudentAddSubjectCode] = useState<string>('')
  const [studentAddSubjectBlock, setStudentAddSubjectBlock] = useState<string>('')
  const [pendingRemovalAssignment, setPendingRemovalAssignment] = useState<string>('')
  const [showProgressiveBalanceDialog, setShowProgressiveBalanceDialog] = useState<boolean>(false)
  const [progressiveBalanceMaxOffset, setProgressiveBalanceMaxOffset] = useState<number>(-4)
  const fileInputRef = useRef<HTMLInputElement | null>(null)

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
    setBalanceResults(null)
    setBalanceHistory([])
    setBalanceMessage('')
    setDebugGroups([])
    setBalanceDeltaCounts(new Map())
    setBalanceBlockDeltaCounts(new Map())
    setErrorMessage('')
    setSelectedStudentsForMassUpdate(new Set())
    setShowMassUpdateDialog(false)
    setMassUpdateTargetSubject('')
    setMassUpdateTargetBlock('')
    setShowAddSubjectDialog(false)
    setAddSubjectTargetCode('')
    setAddSubjectTargetBlock('')
    setShowStudentAddSubjectDialog(false)
    setStudentAddSubjectCode('')
    setStudentAddSubjectBlock('')
    setPendingRemovalAssignment('')

    // Allow selecting the same file again after clearing data.
    if (fileInputRef.current) {
      fileInputRef.current.value = ''
    }
  }

  useEffect(() => {
    saveToLocalStorage(STORAGE_KEYS.parsedData, parsedData)
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
    }
    saveToLocalStorage(STORAGE_KEYS.uiState, stateToPersist)
  }, [selectedStudentId, studentQuery, subjectQuery, blockFilter, viewMode, onlyBlokkfag, showIncompleteBlocks, showOverloadedStudents, showBlockCollisions, showDuplicateSubjects])

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

  const filteredGroupBreakdowns = useMemo(() => {
    if (!parsedData) {
      return []
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
        // Compare blocks using same logic as sortBlocks
        const aIsNumeric = /^Blokk \d+$/u.test(a.block)
        const bIsNumeric = /^Blokk \d+$/u.test(b.block)

        let blockComparison = 0
        if (aIsNumeric && bIsNumeric) {
          const aNum = parseInt(a.block.split(' ')[1], 10)
          const bNum = parseInt(b.block.split(' ')[1], 10)
          blockComparison = aNum - bNum
        } else if (aIsNumeric) {
          blockComparison = -1
        } else if (bIsNumeric) {
          blockComparison = 1
        } else {
          blockComparison = a.block.localeCompare(b.block)
        }

        // If blocks are equal, sort by subject name
        if (blockComparison !== 0) {
          return blockComparison
        }
        return a.subjectName.localeCompare(b.subjectName)
      })
  }, [parsedData, subjectQuery, blockFilter, onlyBlokkfag])

  const groupMemberLookup = useMemo(() => {
    if (!parsedData) {
      return new Map<string, Array<{ id: string; fullName: string }>>()
    }

    const lookup = new Map<string, Map<string, { id: string; fullName: string }>>()
    parsedData.students.forEach((student) => {
      student.assignments.forEach((assignment) => {
        const key = `${assignment.subjectCode}-${assignment.groupCode}-${assignment.block}`
        if (!lookup.has(key)) {
          lookup.set(key, new Map<string, { id: string; fullName: string }>())
        }
        lookup.get(key)?.set(student.id, {
          id: student.id,
          fullName: student.fullName,
        })
      })
    })

    const sortedLookup = new Map<string, Array<{ id: string; fullName: string }>>()
    lookup.forEach((studentMap, key) => {
      const sortedStudents = Array.from(studentMap.values()).sort((a, b) => a.fullName.localeCompare(b.fullName))
      sortedLookup.set(key, sortedStudents)
    })
    return sortedLookup
  }, [parsedData])

  const selectedGroupMembers = useMemo(() => {
    if (!selectedGroupKey) {
      return []
    }
    return groupMemberLookup.get(selectedGroupKey) || []
  }, [groupMemberLookup, selectedGroupKey])

  const summarizedBalanceResults = useMemo(() => {
    if (!balanceResults) {
      return [] as BalanceChange[]
    }

    return summarizeBalanceChanges(balanceResults)
  }, [balanceResults])

  const balanceHistoryByStudent = useMemo(() => {
    const grouped = new Map<string, { studentId: string; studentName: string; classGroup: string; runs: Array<{ runId: string; createdAt: string; message: string; changes: BalanceChange[] }> }>()
    const studentClassById = new Map<string, string>()
    ;(parsedData?.students || []).forEach((student) => {
      studentClassById.set(student.id, student.classGroup || '')
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
        if (!grouped.has(studentId)) {
          grouped.set(studentId, { studentId, studentName, classGroup, runs: [] })
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

  const appendBalanceHistoryRun = (changes: BalanceChange[], message: string): void => {
    const summarized = summarizeBalanceChanges(changes)
    if (summarized.length === 0) {
      return
    }

    const run: BalanceResultRun = {
      id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      createdAt: new Date().toISOString(),
      message,
      changes: summarized,
    }

    setBalanceHistory((previous) => [...previous, run])
  }

  const handleExportBalanceResultsWord = (): void => {
    if (balanceHistoryByStudent.length === 0) {
      setErrorMessage('Ingen balanseringsresultater a eksportere enda.')
      return
    }

    const now = new Date()
    const timestampForFile = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}-${String(now.getHours()).padStart(2, '0')}${String(now.getMinutes()).padStart(2, '0')}`
    const fileName = `balanseringsresultater-${timestampForFile}.doc`

    const bodyRows = balanceHistoryByStudent
      .map((student) => {
        const runHtml = student.runs
          .map((run) => {
            const lines = run.changes
              .map((change) => {
                const subjectName = escapeHtml(change.subjectName || change.subjectCode)
                const fromBlock = escapeHtml(change.fromBlock)
                const fromGroup = escapeHtml(change.fromGroupCode)
                const toBlock = escapeHtml(change.toBlock)
                const toGroup = escapeHtml(change.toGroupCode)
                return `<li><strong>${subjectName}</strong>: ${fromBlock} (${fromGroup}) -> ${toBlock} (${toGroup})</li>`
              })
              .join('')
            return `<div class="run"><div class="run-head"><span>${escapeHtml(formatTimestamp(run.createdAt))}</span></div><ul>${lines}</ul></div>`
          })
          .join('')

        return `<section class="student"><h3>${escapeHtml(student.studentName)} (${escapeHtml(student.studentId)})</h3>${runHtml}</section>`
      })
      .join('')

    const htmlDocument = `<!doctype html><html><head><meta charset="utf-8"><title>Balanseringsresultater</title><style>body{font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#1f2b3d;margin:24px;}h1{font-size:18pt;margin:0 0 12px;}h3{font-size:12pt;margin:14px 0 8px;padding-bottom:4px;border-bottom:1px solid #d9e3f0;}.student{margin-bottom:14px;}.run{padding:6px 0;border-top:1px solid #e7edf6;}.run:first-of-type{border-top:none;}.run-head{display:flex;justify-content:space-between;font-size:9pt;color:#506480;margin-bottom:3px;}ul{margin:0;padding-left:18px;}li{margin:2px 0;}</style></head><body><h1>Balanseringsresultater</h1>${bodyRows}</body></html>`

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
      const sourceName = parsedData.sourceFileName || 'novaschem-eksport.txt'
      const hasTxtExtension = /\.txt$/iu.test(sourceName)
      const baseName = hasTxtExtension ? sourceName.replace(/\.txt$/iu, '') : sourceName
      const downloadName = `${baseName}-endret.txt`

      const blob = new Blob([exportText], { type: 'text/plain;charset=utf-8' })
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
      setParsedData(parsed)
      setSelectedStudentId(parsed.students[0]?.id || '')
      setSelectedGroupKey('')
      setStudentQuery('')
      setSubjectQuery('')
      setBlockFilter('')
      setErrorMessage('')
    } catch {
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
        <input id="novaschem-file" ref={fileInputRef} type="file" accept=".txt" onChange={handleFileUpload} />
        <div className="storage-controls">
          <button type="button" onClick={clearStoredData} className="storage-button danger">
            Tøm lagret data
          </button>
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
            </div>
            <button type="button" className="export-button" onClick={handleExport}>
              Eksporter TXT
            </button>
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
                  onClick={() => setShowIncompleteBlocks(!showIncompleteBlocks)}
                  className={`filter-button ${showIncompleteBlocks ? 'active' : ''}`}
                >
                  Mangler fag ({studentFilterCounts.missingSubjectsCount})
                </button>
                <button
                  type="button"
                  onClick={() => setShowOverloadedStudents(!showOverloadedStudents)}
                  className={`filter-button ${showOverloadedStudents ? 'active' : ''}`}
                >
                  For mange fag ({studentFilterCounts.tooManySubjectsCount})
                </button>
                <button
                  type="button"
                  onClick={() => setShowBlockCollisions(!showBlockCollisions)}
                  className={`filter-button ${showBlockCollisions ? 'active' : ''}`}
                >
                  Blokk-kollisjoner ({studentFilterCounts.blockCollisionsCount})
                </button>
                <button
                  type="button"
                  onClick={() => setShowDuplicateSubjects(!showDuplicateSubjects)}
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
                          <th style={{ width: '80px' }}>Handling</th>
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
                                    const { groupBreakdowns, blockBreakdowns } = recalculateBreakdowns(updatedStudents)

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
          ) : (
            <section className="subject-panel">
              <div className="subject-view-header">
                <h2>Fagvisning</h2>
                <div>
                  <button
                    type="button"
                    onClick={() => setShowAddSubjectDialog(true)}
                    className="balance-button"
                  >
                    Legg til fag
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                    // Calculate original group occupancy
                    const originalCounts = new Map<string, number>()
                    parsedData.groupBreakdowns.forEach((item) => {
                      const key = `${item.subjectCode}|${item.groupCode}|${item.block}`
                      originalCounts.set(key, item.studentCount)
                    })

                    // Calculate original block occupancy
                    const originalBlockCounts = new Map<string, number>()
                    parsedData.blockBreakdowns.forEach((item) => {
                      originalBlockCounts.set(item.block, item.studentCount)
                    })

                    const result = balanceGroups(parsedData.students, (groups) => setDebugGroups(groups))
                    setBalanceResults(result.changes)
                    
                    // Apply changes to the parsed data
                    let updatedData = parsedData
                    if (result.changes.length > 0) {
                      const updatedStudents = applyBalanceChanges(parsedData.students, result.changes)
                      const { groupBreakdowns: newGroupBreakdowns, blockBreakdowns: newBlockBreakdowns } = recalculateBreakdowns(updatedStudents)
                      updatedData = {
                        ...parsedData,
                        students: updatedStudents,
                        groupBreakdowns: newGroupBreakdowns,
                        blockBreakdowns: newBlockBreakdowns,
                      }
                      setParsedData(updatedData)
                    }

                    // Calculate new group occupancy and compute deltas
                    const newCounts = new Map<string, number>()
                    updatedData.groupBreakdowns.forEach((item) => {
                      const key = `${item.subjectCode}|${item.groupCode}|${item.block}`
                      newCounts.set(key, item.studentCount)
                    })

                    const deltas = new Map<string, number>()
                    originalCounts.forEach((originalCount, key) => {
                      const newCount = newCounts.get(key) || 0
                      const delta = newCount - originalCount
                      if (delta !== 0) {
                        deltas.set(key, delta)
                      }
                    })

                    // Calculate new block occupancy and compute block deltas
                    const newBlockCounts = new Map<string, number>()
                    updatedData.blockBreakdowns.forEach((item) => {
                      newBlockCounts.set(item.block, item.studentCount)
                    })

                    const blockDeltas = new Map<string, number>()
                    originalBlockCounts.forEach((originalCount, block) => {
                      const newCount = newBlockCounts.get(block) || 0
                      const delta = newCount - originalCount
                      if (delta !== 0) {
                        blockDeltas.set(block, delta)
                      }
                    })

                    setBalanceDeltaCounts(deltas)
                    setBalanceBlockDeltaCounts(blockDeltas)
                    
                    let nextBalanceMessage = ''
                    if (result.overcrowdedCount === 0) {
                      nextBalanceMessage = '✓ Ingen overfulle grupper. Alle grupper er innenfor kapasitet.'
                    } else if (result.changes.length === 0) {
                      nextBalanceMessage = `Fant ${result.overcrowdedCount} overfull(e) gruppe(r), men ingen elever kan flyttes (alle elever har unike blokktildelinger i disse fagene).`
                    } else {
                      const uniqueStudents = new Set(result.changes.map((c) => c.studentId))
                      nextBalanceMessage = `Fant ${result.overcrowdedCount} overfull(e) gruppe(r). Flyttet ${uniqueStudents.size} elev(er) (${result.changes.length} fagendringer).`
                      appendBalanceHistoryRun(result.changes, nextBalanceMessage)
                    }
                    setBalanceMessage(nextBalanceMessage)
                    }}
                    className="balance-button"
                  >
                    Balanser
                  </button>
                  <button
                    type="button"
                    onClick={() => setShowProgressiveBalanceDialog(true)}
                    className="balance-button"
                  >
                    Progressiv balansering
                  </button>
                </div>
              </div>

              {(balanceResults !== null || balanceHistoryByStudent.length > 0) && (
                <div className="balance-results">
                  <h3>
                    Balanseringsresultater{' '}
                    {balanceHistoryByStudent.length > 0
                      ? `(${balanceHistoryByStudent.length} elev${balanceHistoryByStudent.length === 1 ? '' : 'er'} med historikk)`
                      : ''}
                  </h3>
                  {balanceMessage && <p className="balance-message">{balanceMessage}</p>}
                  {balanceHistoryByStudent.length > 0 && (
                    <div className="balance-list">
                      {balanceHistoryByStudent.map((student) => (
                        <div key={student.studentId} className="balance-item">
                          <strong>{student.studentName}</strong> ({student.classGroup || '-'}) ({student.studentId})
                          {student.runs.map((run) => (
                            <div key={run.runId} className="balance-run-segment">
                              <div className="balance-run-meta">
                                <span>{run.message || 'Balansering'}</span>
                                <span>{formatTimestamp(run.createdAt)}</span>
                              </div>
                              {run.changes.map((change) => (
                                <div key={`${run.runId}-${change.subjectCode}-${change.fromGroupCode}-${change.toGroupCode}-${change.fromBlock}-${change.toBlock}`} className="balance-change-line balance-change-moved">
                                  <strong>{change.subjectName || change.subjectCode}</strong>: {change.fromBlock} ({change.fromGroupCode}) &rarr; {change.toBlock} ({change.toGroupCode})
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
                      setBalanceDeltaCounts(new Map())
                      setBalanceBlockDeltaCounts(new Map())
                    }}
                    className="clear-results-button"
                  >
                    Tøm
                  </button>
                </div>
              )}

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

              <h2>Per faggruppe</h2>
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
                  {filteredGroupBreakdowns.map((item) => {
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
                                  Klikk en annen rad for å bytte, eller klikk igjen for å lukke.
                                </p>
                              </div>
                              <div style={{ marginBottom: '1rem', display: 'flex', gap: '0.5rem' }}>
                                <button
                                  type="button"
                                  onClick={(e) => {
                                    e.stopPropagation()
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
                                  <div key={student.id} className="group-member-item" style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                                    <input
                                      type="checkbox"
                                      checked={selectedStudentsForMassUpdate.has(student.id)}
                                      onChange={(e) => {
                                        e.stopPropagation()
                                        const newSet = new Set(selectedStudentsForMassUpdate)
                                        if (e.target.checked) {
                                          newSet.add(student.id)
                                        } else {
                                          newSet.delete(student.id)
                                        }
                                        setSelectedStudentsForMassUpdate(newSet)
                                      }}
                                      onClick={(e) => e.stopPropagation()}
                                    />
                                    <div style={{ flex: 1 }}>
                                      <span>{student.fullName}</span>
                                      <small>{student.id}</small>
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
                <div className="modal-overlay" onClick={() => setShowMassUpdateDialog(false)}>
                  <div className="modal-content" onClick={(e) => e.stopPropagation()}>
                    <h3>Massoppdater elever</h3>
                    <p>{selectedStudentsForMassUpdate.size} elever valgt</p>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '1rem', marginTop: '1rem' }}>
                      <label>
                        <strong>Nytt fag:</strong>
                        <select
                          value={massUpdateTargetSubject}
                          onChange={(e) => setMassUpdateTargetSubject(e.target.value)}
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
                          <option value="Blokk 1">Blokk 1</option>
                          <option value="Blokk 2">Blokk 2</option>
                          <option value="Blokk 3">Blokk 3</option>
                          <option value="Blokk 4">Blokk 4</option>
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

                            // Store original counts
                            const originalCounts = new Map<string, number>()
                            parsedData!.groupBreakdowns.forEach((item) => {
                              const key = `${item.subjectCode}|${item.groupCode}|${item.block}`
                              originalCounts.set(key, item.studentCount)
                            })

                            const originalBlockCounts = new Map<string, number>()
                            parsedData!.blockBreakdowns.forEach((item) => {
                              originalBlockCounts.set(item.block, item.studentCount)
                            })

                            // Apply mass update by directly manipulating student assignments
                            const updatedStudents = parsedData!.students.map((student) => {
                              if (!selectedStudentsForMassUpdate.has(student.id)) {
                                return student
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
                            const { groupBreakdowns: newGroupBreakdowns, blockBreakdowns: newBlockBreakdowns } = recalculateBreakdowns(updatedStudents)
                            const updatedData = {
                              ...parsedData!,
                              students: updatedStudents,
                              groupBreakdowns: newGroupBreakdowns,
                              blockBreakdowns: newBlockBreakdowns,
                            }
                            setParsedData(updatedData)

                            // Calculate deltas
                            const newCounts = new Map<string, number>()
                            updatedData.groupBreakdowns.forEach((item) => {
                              const key = `${item.subjectCode}|${item.groupCode}|${item.block}`
                              newCounts.set(key, item.studentCount)
                            })

                            const allGroupKeys = new Set<string>([...Array.from(originalCounts.keys()), ...Array.from(newCounts.keys())])
                            const deltas = new Map<string, number>()
                            allGroupKeys.forEach((key) => {
                              const delta = (newCounts.get(key) || 0) - (originalCounts.get(key) || 0)
                              if (delta !== 0) {
                                deltas.set(key, delta)
                              }
                            })

                            const newBlockCounts = new Map<string, number>()
                            updatedData.blockBreakdowns.forEach((item) => {
                              newBlockCounts.set(item.block, item.studentCount)
                            })

                            const blockDeltas = new Map<string, number>()
                            originalBlockCounts.forEach((originalCount, block) => {
                              const newCount = newBlockCounts.get(block) || 0
                              const delta = newCount - originalCount
                              if (delta !== 0) {
                                blockDeltas.set(block, delta)
                              }
                            })

                            setBalanceDeltaCounts(deltas)
                            setBalanceBlockDeltaCounts(blockDeltas)

                            // Update message
                            setBalanceMessage(`Masseoppdatering: ${selectedStudentsForMassUpdate.size} elever flyttet til ${targetSubject.name} (${smallestGroup.groupCode}) i ${massUpdateTargetBlock}`)

                            // Clear selection and close dialog
                            setSelectedStudentsForMassUpdate(new Set())
                            setShowMassUpdateDialog(false)
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
                            setShowMassUpdateDialog(false)
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
                    <p>Legg til fag som tilgjengelig i valgt blokk (uten å tildele elever).</p>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '1rem', marginTop: '1rem' }}>
                      <label>
                        <strong>Fag:</strong>
                        <select
                          value={addSubjectTargetCode}
                          onChange={(e) => setAddSubjectTargetCode(e.target.value)}
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
                        <strong>Blokk:</strong>
                        <select
                          value={addSubjectTargetBlock}
                          onChange={(e) => setAddSubjectTargetBlock(e.target.value)}
                          style={{ width: '100%', marginTop: '0.5rem', padding: '0.5rem' }}
                        >
                          <option value="">Velg blokk...</option>
                          <option value="Blokk 1">Blokk 1</option>
                          <option value="Blokk 2">Blokk 2</option>
                          <option value="Blokk 3">Blokk 3</option>
                          <option value="Blokk 4">Blokk 4</option>
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

                            const updatedSubjects = parsedData.subjects.map((subject) => {
                              if (subject.code !== addSubjectTargetCode) {
                                return subject
                              }

                              const nextBlocks = subject.blocks.includes(addSubjectTargetBlock)
                                ? subject.blocks
                                : sortBlocks([...subject.blocks, addSubjectTargetBlock])

                              return {
                                ...subject,
                                blocks: nextBlocks,
                              }
                            })

                            const updatedBlocks = parsedData.blocks.includes(addSubjectTargetBlock)
                              ? parsedData.blocks
                              : sortBlocks([...parsedData.blocks, addSubjectTargetBlock])

                            const hasGroupInTargetBlock = parsedData.groupBreakdowns.some(
                              (group) => group.subjectCode === addSubjectTargetCode && group.block === addSubjectTargetBlock
                            )

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

                            const wasAlreadyAvailable = targetSubject.blocks.includes(addSubjectTargetBlock)
                            if (wasAlreadyAvailable) {
                              setBalanceMessage(`${targetSubject.name} er allerede tilgjengelig i ${addSubjectTargetBlock}.`)
                            } else {
                              setBalanceMessage(`${targetSubject.name} er nå tilgjengelig i ${addSubjectTargetBlock}.`)
                            }

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

                            // Calculate original group occupancy
                            const originalCounts = new Map<string, number>()
                            parsedData.groupBreakdowns.forEach((item) => {
                              const key = `${item.subjectCode}|${item.groupCode}|${item.block}`
                              originalCounts.set(key, item.studentCount)
                            })

                            // Calculate original block occupancy
                            const originalBlockCounts = new Map<string, number>()
                            parsedData.blockBreakdowns.forEach((item) => {
                              originalBlockCounts.set(item.block, item.studentCount)
                            })

                            const result = progressiveBalanceGroups(parsedData.students, progressiveBalanceMaxOffset, (groups) => setDebugGroups(groups))
                            setBalanceResults(result.allChanges)
                            
                            // Apply changes to the parsed data
                            let updatedData = parsedData
                            if (result.allChanges.length > 0) {
                              const updatedStudents = applyBalanceChanges(parsedData.students, result.allChanges)
                              const { groupBreakdowns: newGroupBreakdowns, blockBreakdowns: newBlockBreakdowns } = recalculateBreakdowns(updatedStudents)
                              updatedData = {
                                ...parsedData,
                                students: updatedStudents,
                                groupBreakdowns: newGroupBreakdowns,
                                blockBreakdowns: newBlockBreakdowns,
                              }
                              setParsedData(updatedData)
                            }

                            // Calculate new group occupancy and compute deltas
                            const newCounts = new Map<string, number>()
                            updatedData.groupBreakdowns.forEach((item) => {
                              const key = `${item.subjectCode}|${item.groupCode}|${item.block}`
                              newCounts.set(key, item.studentCount)
                            })

                            const deltas = new Map<string, number>()
                            originalCounts.forEach((originalCount, key) => {
                              const newCount = newCounts.get(key) || 0
                              const delta = newCount - originalCount
                              if (delta !== 0) {
                                deltas.set(key, delta)
                              }
                            })

                            // Calculate new block occupancy and compute block deltas
                            const newBlockCounts = new Map<string, number>()
                            updatedData.blockBreakdowns.forEach((item) => {
                              newBlockCounts.set(item.block, item.studentCount)
                            })

                            const blockDeltas = new Map<string, number>()
                            originalBlockCounts.forEach((originalCount, block) => {
                              const newCount = newBlockCounts.get(block) || 0
                              const delta = newCount - originalCount
                              if (delta !== 0) {
                                blockDeltas.set(block, delta)
                              }
                            })

                            setBalanceDeltaCounts(deltas)
                            setBalanceBlockDeltaCounts(blockDeltas)
                            
                            setBalanceMessage(result.summary)
                            appendBalanceHistoryRun(result.allChanges, result.summary)
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



              <h2>Per fag (alle grupper)</h2>
              <table>
                <thead>
                  <tr>
                    <th>Fag</th>
                    <th>Tittel</th>
                    <th>Elever</th>
                    <th>Blokker</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredSubjects.map((subject) => (
                    <tr key={subject.code}>
                      <td>{subject.code}</td>
                      <td>{subject.name}</td>
                      <td>{subject.studentCount}</td>
                      <td>{subject.blocks.length > 0 ? subject.blocks.join(', ') : '-'}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
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
                      <strong>Velg blokk og gruppe:</strong>
                      <div style={{ marginTop: '0.5rem', display: 'flex', flexDirection: 'column', gap: '0.5rem' }}>
                        {parsedData.groupBreakdowns
                          .filter((group) => group.subjectCode === studentAddSubjectCode)
                          .sort((a, b) => {
                            const blockCompare = sortBlocks([a.block, b.block]).indexOf(a.block) - sortBlocks([a.block, b.block]).indexOf(a.block)
                            if (blockCompare !== 0) return blockCompare
                            return a.groupCode.localeCompare(b.groupCode)
                          })
                          .map((group) => (
                            <button
                              key={`${group.groupCode}-${group.block}`}
                              type="button"
                              onClick={() => setStudentAddSubjectBlock(`${group.groupCode}|${group.block}`)}
                              className={studentAddSubjectBlock === `${group.groupCode}|${group.block}` ? 'filter-button active' : 'filter-button'}
                              style={{
                                padding: '0.75rem',
                                textAlign: 'left',
                                fontWeight: 'normal',
                                backgroundColor: studentAddSubjectBlock === `${group.groupCode}|${group.block}` ? '#0969da' : '#f9fafb',
                                color: studentAddSubjectBlock === `${group.groupCode}|${group.block}` ? '#ffffff' : '#24292f',
                                border: studentAddSubjectBlock === `${group.groupCode}|${group.block}` ? '1px solid #0969da' : '1px solid #d0d7de',
                              }}
                            >
                              {group.block} - {group.subjectName} ({group.studentCount} elever)
                            </button>
                          ))}
                      </div>
                    </div>
                  )}

                  <div style={{ display: 'flex', gap: '0.5rem', marginTop: '0.5rem' }}>
                    <button
                      type="button"
                      onClick={() => {
                        if (!studentAddSubjectCode || !studentAddSubjectBlock) {
                          alert('Vennligst velg fag og gruppe')
                          return
                        }

                        const [groupCode, block] = studentAddSubjectBlock.split('|')
                        const targetSubject = parsedData.subjects.find((s) => s.code === studentAddSubjectCode)
                        if (!targetSubject) {
                          alert('Fag ikke funnet')
                          return
                        }

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
                        const { groupBreakdowns, blockBreakdowns } = recalculateBreakdowns(updatedStudents)

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
        </>
      )}
    </div>
  )
}

export default App
