import { useEffect, useMemo, useState } from 'react'
import './App.css'

type RawTable = {
  columns: string[]
  rows: string[][]
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
        if (otherSubject === subjectCode && !blockSet.has(otherBlock) && occupants.length < maxCap) {
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
  // Create a deep copy of students with modified assignments
  return students.map((student) => {
    const studentChanges = changes.filter((c) => c.studentId === student.id)
    if (studentChanges.length === 0) {
      return student
    }

    // Apply each change to this student's assignments
    const updatedAssignments = student.assignments.map((assignment) => {
      const change = studentChanges.find(
        (c) => c.subjectCode === assignment.subjectCode && c.fromGroupCode === assignment.groupCode && c.fromBlock === assignment.block
      )
      if (change) {
        // Move this assignment to the new group and block
        return {
          ...assignment,
          groupCode: change.toGroupCode,
          block: change.toBlock,
        }
      }
      return assignment
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

function extractTables(rawText: string): Map<string, RawTable> {
  const lines = rawText.replace(/\r\n/gu, '\n').split('\n')
  const tables = new Map<string, RawTable>()
  const tableHeaderRegex = /^([A-Za-z_]+)\s+\(\d+\)$/u

  let index = 0
  while (index < lines.length) {
    const sectionLine = lines[index].trim()
    const sectionMatch = sectionLine.match(tableHeaderRegex)
    if (!sectionMatch) {
      index += 1
      continue
    }

    const tableName = sectionMatch[1]
    index += 1
    while (index < lines.length && lines[index].trim() !== '[Rows]') {
      index += 1
    }
    if (index >= lines.length) {
      break
    }

    index += 1
    while (index < lines.length && lines[index].trim() === '') {
      index += 1
    }
    if (index >= lines.length) {
      break
    }

    const rawHeaderCells = lines[index].split('\t').map((cell) => sanitizeHeader(cell))
    const columns = rawHeaderCells.filter((header) => header.length > 0)
    index += 1

    const rows: string[][] = []
    while (index < lines.length) {
      const rawLine = lines[index]
      const trimmedLine = rawLine.trim()

      if (trimmedLine === '') {
        index += 1
        break
      }
      if (tableHeaderRegex.test(trimmedLine)) {
        break
      }

      const cells = rawLine.split('\t')
      if (cells.some((cell) => cell.trim() !== '')) {
        while (cells.length < columns.length) {
          cells.push('')
        }
        rows.push(cells.slice(0, columns.length))
      }
      index += 1
    }

    tables.set(tableName, { columns, rows })
  }

  return tables
}

function rowsToObjects(table?: RawTable): Array<Record<string, string>> {
  if (!table) {
    return []
  }

  return table.rows.map((row) => {
    const objectRow: Record<string, string> = {}
    table.columns.forEach((column, i) => {
      objectRow[column] = (row[i] || '').trim()
    })
    return objectRow
  })
}

function parseNovaschemExport(rawText: string): ParsedData {
  const tables = extractTables(rawText)
  const tableNames = Array.from(tables.keys()).sort((a, b) => a.localeCompare(b))

  const studentRows = rowsToObjects(tables.get('Student'))
  const subjectRows = rowsToObjects(tables.get('Subject'))
  const groupRows = rowsToObjects(tables.get('Group'))
  const groupStudentRows = rowsToObjects(tables.get('Group_Student'))
  const taRows = rowsToObjects(tables.get('TA'))

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

  return {
    students,
    subjects,
    groupBreakdowns,
    blockBreakdowns,
    blocks,
    tableNames,
  }
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

  const [parsedData, setParsedData] = useState<ParsedData | null>(() =>
    loadFromLocalStorage<ParsedData | null>(STORAGE_KEYS.parsedData, null)
  )
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

  const clearStoredData = (): void => {
    removeFromLocalStorage(STORAGE_KEYS.parsedData)
    removeFromLocalStorage(STORAGE_KEYS.uiState)

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
  }

  useEffect(() => {
    saveToLocalStorage(STORAGE_KEYS.parsedData, parsedData)
  }, [parsedData])

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

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
    const file = event.target.files?.[0]
    if (!file) {
      return
    }

    try {
      const fileBuffer = await file.arrayBuffer()
      const text = decodeBestEffort(fileBuffer)
      const parsed = parseNovaschemExport(text)
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
        <input id="novaschem-file" type="file" accept=".txt" onChange={handleFileUpload} />
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
                    
                    if (result.overcrowdedCount === 0) {
                      setBalanceMessage('✓ Ingen overfulle grupper. Alle grupper er innenfor kapasitet.')
                    } else if (result.changes.length === 0) {
                      setBalanceMessage(`Fant ${result.overcrowdedCount} overfull(e) gruppe(r), men ingen elever kan flyttes (alle elever har unike blokktildelinger i disse fagene).`)
                    } else {
                      const uniqueStudents = new Set(result.changes.map((c) => c.studentId))
                      setBalanceMessage(`Fant ${result.overcrowdedCount} overfull(e) gruppe(r). Flyttet ${uniqueStudents.size} elev(er) (${result.changes.length} fagendringer).`)
                    }
                    }}
                    className="balance-button"
                  >
                    Balanser
                  </button>
                </div>
              </div>

              {balanceResults !== null && (
                <div className="balance-results">
                  <h3>
                    Balanseringsresultater{' '}
                    {balanceResults.length > 0
                      ? `(${new Set(balanceResults.map((c) => c.studentId)).size} elev${new Set(balanceResults.map((c) => c.studentId)).size === 1 ? '' : 'er'}, ${balanceResults.length} fagendringer)`
                      : ''}
                  </h3>
                  {balanceMessage && <p className="balance-message">{balanceMessage}</p>}
                  {balanceResults.length > 0 && (
                    <div className="balance-list">
                      {(() => {
                        const groupedByStudent = new Map<string, BalanceChange[]>()
                        balanceResults.forEach((change) => {
                          if (!groupedByStudent.has(change.studentId)) {
                            groupedByStudent.set(change.studentId, [])
                          }
                          groupedByStudent.get(change.studentId)?.push(change)
                        })

                        return Array.from(groupedByStudent.values()).map((changes) => (
                          <div key={changes[0].studentId} className="balance-item">
                            <strong>{changes[0].studentName}</strong> ({changes[0].studentId})
                            <br />
                            {changes.map((change, i) => (
                              <div key={i} style={{ marginLeft: '1rem', marginTop: '0.3rem', fontSize: '0.95rem' }}>
                                <strong>{change.subjectCode}</strong> ({change.subjectName}): Gruppe {change.fromGroupCode}, {change.fromBlock} → Gruppe {change.toGroupCode}, {change.toBlock}
                              </div>
                            ))}
                          </div>
                        ))
                      })()}
                    </div>
                  )}
                  {balanceResults.length === 0 && debugGroups.length > 0 && (
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
                    onClick={() => {
                      setBalanceResults(null)
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
