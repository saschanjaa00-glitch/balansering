import { useMemo, useState } from 'react'
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

function getMaxCapacityForSubject(subjectName: string): number {
  return SUBJECT_MAX_CAPACITY[subjectName] || 30
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
      .find((a) => a.subjectCode === subjectCode)?.subjectName || 'UNKNOWN'

    const maxCap = getMaxCapacityForSubject(subjectName)
    const status = studentIds.length > maxCap ? `OVERCROWDED (${studentIds.length} > ${maxCap})` : `OK (${studentIds.length} ≤ ${maxCap})`
    
    debugGroups.push({ key, count: studentIds.length, maxCap, status })

    if (subjectName !== 'UNKNOWN' && studentIds.length > maxCap) {
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
  const [parsedData, setParsedData] = useState<ParsedData | null>(null)
  const [selectedStudentId, setSelectedStudentId] = useState<string>('')
  const [studentQuery, setStudentQuery] = useState<string>('')
  const [subjectQuery, setSubjectQuery] = useState<string>('')
  const [blockFilter, setBlockFilter] = useState<string>('')
  const [viewMode, setViewMode] = useState<'students' | 'subjects'>('students')
  const [selectedGroupKey, setSelectedGroupKey] = useState<string>('')
  const [onlyBlokkfag, setOnlyBlokkfag] = useState<boolean>(true)
  const [showIncompleteBlocks, setShowIncompleteBlocks] = useState<boolean>(false)
  const [balanceResults, setBalanceResults] = useState<BalanceChange[] | null>(null)
  const [balanceMessage, setBalanceMessage] = useState<string>('')
  const [debugGroups, setDebugGroups] = useState<Array<{ key: string; count: number; maxCap: number; status: string }>>([])
  const [balanceDeltaCounts, setBalanceDeltaCounts] = useState<Map<string, number>>(new Map())
  const [balanceBlockDeltaCounts, setBalanceBlockDeltaCounts] = useState<Map<string, number>>(new Map())
  const [errorMessage, setErrorMessage] = useState<string>('')

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
        const blocksWithSubjects = new Set<string>()
        student.assignments.forEach((assignment) => {
          if (/^Blokk [1-4]$/.test(assignment.block)) {
            blocksWithSubjects.add(assignment.block)
          }
        })
        return blocksWithSubjects.size < 3
      }
      return true
    })
  }, [parsedData, studentQuery, showIncompleteBlocks])

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

    return filtered.sort((a, b) => {
      const aIsMatte = a.block === 'MATTE'
      const bIsMatte = b.block === 'MATTE'
      if (aIsMatte && !bIsMatte) return 1
      if (!aIsMatte && bIsMatte) return -1
      return 0
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
        const matchesBlock = !blockFilter || item.block === blockFilter
        const isBlokkfag = item.block && item.block.trim().length > 0
        return matchesSearch && matchesBlock && (!onlyBlokkfag || isBlokkfag)
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
      setErrorMessage('Could not parse the export file. Confirm that the file is a Novaschem TXT export.')
    }
  }

  return (
    <div className="app-shell">
      <header className="hero">
        <p className="hero-tag">Novaschem to SATS</p>
        <h1>Timetable Export Viewer</h1>
        <p className="hero-subtitle">
          Upload one Novaschem TXT export to browse students, their selected subjects, and assigned blocks.
        </p>
      </header>

      <section className="upload-panel">
        <label htmlFor="novaschem-file" className="file-label">
          Select Novaschem TXT File
        </label>
        <input id="novaschem-file" type="file" accept=".txt" onChange={handleFileUpload} />
        {errorMessage && <p className="error-text">{errorMessage}</p>}
        {parsedData && (
          <div className="stats-grid">
            <article>
              <strong>{parsedData.students.length}</strong>
              <span>Students</span>
            </article>
            <article>
              <strong>{parsedData.subjects.length}</strong>
              <span>Subjects In Use</span>
            </article>
            <article>
              <strong>{parsedData.blocks.length}</strong>
              <span>Blocks Found</span>
            </article>
            <article>
              <strong>{parsedData.tableNames.length}</strong>
              <span>Tables Parsed</span>
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
                Student View
              </button>
              <button
                type="button"
                className={viewMode === 'subjects' ? 'active' : ''}
                onClick={() => setViewMode('subjects')}
              >
                Subject View
              </button>
            </div>

            <input
              type="search"
              value={subjectQuery}
              onChange={(event) => setSubjectQuery(event.target.value)}
              placeholder="Filter by subject code or name"
            />

            <label className="blokkfag-checkbox">
              <input
                type="checkbox"
                checked={onlyBlokkfag}
                onChange={(event) => setOnlyBlokkfag(event.target.checked)}
              />
              <span>Blokkfag only</span>
            </label>

            <select value={blockFilter} onChange={(event) => setBlockFilter(event.target.value)}>
              <option value="">All blocks</option>
              {parsedData.blocks.map((block) => (
                <option key={block} value={block}>
                  {block}
                </option>
              ))}
            </select>
          </section>

          {viewMode === 'students' ? (
            <section className="viewer-grid">
              <aside className="student-list-panel">
                <input
                  type="search"
                  value={studentQuery}
                  onChange={(event) => setStudentQuery(event.target.value)}
                  placeholder="Search student name or number"
                />

                <button
                  type="button"
                  onClick={() => setShowIncompleteBlocks(!showIncompleteBlocks)}
                  className={`filter-button ${showIncompleteBlocks ? 'active' : ''}`}
                >
                  Show only missing 2+ blocks
                </button>

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
                        {student.id} | {student.assignments.length} subjects
                      </small>
                    </button>
                  ))}
                </div>
              </aside>

              <article className="detail-panel">
                {selectedStudent ? (
                  <>
                    <h2>
                      {selectedStudent.fullName} {selectedStudent.classGroup && `(${selectedStudent.classGroup})`}
                    </h2>
                    <p>
                      Student number: <strong>{selectedStudent.id}</strong>
                      {selectedStudent.email ? <> | {selectedStudent.email}</> : null}
                    </p>

                    <table>
                      <thead>
                        <tr>
                          <th>Subject</th>
                          <th>Title</th>
                          <th>Group</th>
                          <th>Block</th>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredAssignments.map((assignment) => (
                          <tr
                            key={`${assignment.subjectCode}-${assignment.groupCode}-${assignment.block}`}
                            className={assignment.block === 'MATTE' ? 'matte-row' : ''}
                          >
                            <td>{assignment.subjectCode}</td>
                            <td>{assignment.subjectName}</td>
                            <td>{assignment.groupCode}</td>
                            <td>{assignment.block || '-'}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </>
                ) : (
                  <p>Select a student to see subject choices.</p>
                )}
              </article>
            </section>
          ) : (
            <section className="subject-panel">
              <div className="subject-view-header">
                <h2>Subject View</h2>
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
                      setBalanceMessage('✓ No overcrowded groups. All groups are within capacity.')
                    } else if (result.changes.length === 0) {
                      setBalanceMessage(`Found ${result.overcrowdedCount} overcrowded group(s), but no students can be moved (all students have unique block assignments in those subjects).`)
                    } else {
                      const uniqueStudents = new Set(result.changes.map((c) => c.studentId))
                      setBalanceMessage(`Found ${result.overcrowdedCount} overcrowded group(s). Successfully moved ${uniqueStudents.size} student(s) (${result.changes.length} subject changes).`)
                    }
                  }}
                  className="balance-button"
                >
                  Balanser
                </button>
              </div>

              {balanceResults !== null && (
                <div className="balance-results">
                  <h3>
                    Balance Results{' '}
                    {balanceResults.length > 0
                      ? `(${new Set(balanceResults.map((c) => c.studentId)).size} student${new Set(balanceResults.map((c) => c.studentId)).size === 1 ? '' : 's'}, ${balanceResults.length} subject changes)`
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
                                <strong>{change.subjectCode}</strong> ({change.subjectName}): Group {change.fromGroupCode}, {change.fromBlock} → Group {change.toGroupCode}, {change.toBlock}
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
                        All groups analyzed ({debugGroups.length} groups):
                      </p>
                      <div className="balance-list" style={{ maxHeight: '400px' }}>
                        {debugGroups.map((group, idx) => (
                          <div key={idx} className="balance-item" style={{ fontSize: '0.85rem' }}>
                            <strong>{group.key}</strong>
                            <br />
                            Count: {group.count}, Max: {group.maxCap}
                            <br />
                            <span style={{ color: group.status.includes('OVERCROWDED') ? '#c41e3a' : '#2f7044' }}>
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
                    Clear
                  </button>
                </div>
              )}

              <h2>By Block</h2>
              <div className="block-summary-grid">
                {parsedData.blockBreakdowns
                  .filter((item) => !blockFilter || item.block === blockFilter)
                  .map((item) => {
                    const blockDelta = balanceBlockDeltaCounts.get(item.block)
                    return (
                      <article key={item.block} className={item.block === 'MATTE' ? 'matte-block' : ''}>
                        <strong>{item.block}</strong>
                        <span>{item.subjectCount} subjects</span>
                        <span>{item.groupCount} groups</span>
                        <span>
                          {item.studentCount} students
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

              <h2>By Subject Group</h2>
              <table>
                <thead>
                  <tr>
                    <th>Subject</th>
                    <th>Title</th>
                    <th>Group</th>
                    <th>Students</th>
                    <th>Block</th>
                    {balanceDeltaCounts.size > 0 && <th>Change</th>}
                  </tr>
                </thead>
                <tbody>
                  {filteredGroupBreakdowns.map((item) => {
                    const groupKey = `${item.subjectCode}|${item.groupCode}|${item.block}`
                    const delta = balanceDeltaCounts.get(groupKey)
                    return (
                      <tr
                        key={`${item.subjectCode}-${item.groupCode}-${item.block}`}
                        onClick={() => setSelectedGroupKey(`${item.subjectCode}-${item.groupCode}-${item.block}`)}
                        className={
                          selectedGroupKey === `${item.subjectCode}-${item.groupCode}-${item.block}`
                            ? 'clickable-row active'
                            : 'clickable-row'
                        }
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
                    )
                  })}
                </tbody>
              </table>

              {selectedGroupKey && (
                <section className="group-members-panel">
                  <h3>Students In Selected Subject Group</h3>
                  <p>Click another row to switch subject/group/block.</p>
                  <div className="group-member-list">
                    {selectedGroupMembers.map((student) => (
                      <div key={student.id} className="group-member-item">
                        <span>{student.fullName}</span>
                        <small>{student.id}</small>
                      </div>
                    ))}
                  </div>
                </section>
              )}

              <h2>By Subject (All Groups)</h2>
              <table>
                <thead>
                  <tr>
                    <th>Subject</th>
                    <th>Title</th>
                    <th>Students</th>
                    <th>Blocks</th>
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
        </>
      )}
    </div>
  )
}

export default App
