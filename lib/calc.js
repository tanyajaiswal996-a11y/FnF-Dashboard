// FnF calculation logic — implements PRD §8 Business Rules & Calculation Formulas.
// Open Items (PRD §10): TDS methodology and Incentive/PLI treatment are not finalized.
// Both are carried as manual input fields until stakeholder sign-off defines the formula.

function calc(rec) {
  const c = rec.comp || {};
  const d = rec.dues || {};

  const grossSalary = Number(c.grossSalary || 0);
  const epf = Number(c.epf || 0);
  const workingDays = Number(c.workingDays || 0);
  const presentDays = Number(c.presentDays || 0);
  const tenureYears = Number(c.tenureYears || 0);

  // FR-3.2: Salary = Gross Salary ÷ Working Days × Present Days
  const salary = workingDays > 0 ? (grossSalary / workingDays) * presentDays : 0;

  // FR-3.3: PF = EPF# ÷ Working Days × Present Days
  const pf = workingDays > 0 ? (epf / workingDays) * presentDays : 0;

  // FR-3.4: Gratuity only if tenure > 5 years = 50% of Gross Salary × 15 × Tenure ÷ 26
  const gratuity = tenureYears > 5 ? (0.5 * grossSalary * 15 * tenureYears) / 26 : 0;

  // FR-3.5: Notice Pay as recorded in No Dues Panel (HR → Notice Pay sub-department); PF excluded
  const noticePay = d.noticePay && d.noticePay.applicable ? Number(d.noticePay.amount || 0) : 0;

  // FR-2.2: Exam Dues — deduct only if "Yes Required"
  const examDeduct = d.exam && d.exam.required === 'yes_required' ? Number(d.exam.amount || 0) : 0;

  // FR-2.3: WFH Dues — deduct only if balance is negative
  const wfhBalance = Number((d.wfh && d.wfh.balance) || 0);
  const wfhDeduct = wfhBalance < 0 ? Math.abs(wfhBalance) : 0;

  // FR-3.9: Loan & Advance — outstanding recovery
  const loanAdv = Number((d.loanAdvance && d.loanAdvance.amount) || 0);

  // FR-3.10: TA Bill reconciliation — positive = owed to employee, negative = recoverable
  const taBill = Number(c.taBill || 0);

  // FR-3.11: NPS / VPF entries in Net Payable Panel
  const nps = Number(c.nps || 0);
  const vpf = Number(c.vpf || 0);

  // FR-3.6 / FR-3.7: TDS and Incentive/PLI — Open Items §10, manual input until finalized
  const tds = Number(c.tds || 0);
  const pli = Number(c.pli || 0);

  const earnings = salary + pf + gratuity + noticePay + pli + (taBill > 0 ? taBill : 0);
  const deductions = examDeduct + wfhDeduct + loanAdv + tds + nps + vpf + (taBill < 0 ? Math.abs(taBill) : 0);
  const net = earnings - deductions;

  return { salary, pf, gratuity, noticePay, examDeduct, wfhDeduct, loanAdv, taBill, nps, vpf, tds, pli, earnings, deductions, net };
}

function recStatus(rec) {
  if (rec.closure && rec.closure.expLetter && rec.closure.slips) return 'closed';
  if (rec.started) return 'active';
  return 'pending';
}

function daysBetween(a, b) {
  return Math.round((new Date(b) - new Date(a)) / 86400000);
}

// Success Criterion §12 / process note: FnF sheet should reach the employee within 2 days of LWD.
function recSla(rec, asOfDate) {
  if (!rec.lwd) return { state: 'na', label: '—' };
  if (rec.emailSentDate) {
    const gap = daysBetween(rec.lwd, rec.emailSentDate);
    return gap <= 2 ? { state: 'within', label: `Sent day ${gap}` } : { state: 'breached', label: `Sent day ${gap} (late)` };
  }
  const elapsed = daysBetween(rec.lwd, asOfDate);
  if (elapsed < 0) return { state: 'upcoming', label: `LWD in ${-elapsed}d` };
  if (elapsed > 2) return { state: 'breached', label: `${elapsed}d overdue — not sent` };
  return { state: 'waiting', label: `Day ${elapsed} of 2 — not sent yet` };
}

module.exports = { calc, recStatus, recSla, daysBetween };
