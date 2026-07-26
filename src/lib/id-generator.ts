export function generateCertificateId(eventCategory: string = 'ML', index: number = 1): string {
  const year = new Date().getFullYear();
  const categoryCode = eventCategory.toUpperCase().substring(0, 3).replace(/[^A-Z]/g, 'EV');
  const seq = String(index).padStart(3, '0');
  const randomSuffix = Math.random().toString(36).substring(2, 6).toUpperCase();
  return `CSEA-${year}-${categoryCode}-${seq}${randomSuffix}`;
}

export function extractStudentId(row: Record<string, any>): string {
  // Common column names for student id / roll no
  const candidateKeys = ['RollNumber', 'Roll Number ', 'Roll Number', 'Register No', 'Reg No', 'Student ID', 'Roll No', 'RollNo', 'Roll_Number', 'Roll_No', 'Register_No'];
  for (const key of candidateKeys) {
    if (row[key] && String(row[key]).trim() !== '') {
      return String(row[key]).trim();
    }
  }

  // Search any key containing 'roll' or 'reg'
  for (const [k, v] of Object.entries(row)) {
    if (/(roll|reg|id)/i.test(k) && v) {
      return String(v).trim();
    }
  }

  return `STU-${Math.floor(100000 + Math.random() * 900000)}`;
}
