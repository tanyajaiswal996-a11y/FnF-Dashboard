// FnF calculation logic — implements PRD §8 Business Rules & Calculation Formulas.
// Open Items (PRD §10): TDS methodology and Incentive/PLI treatment are not finalized.
// Both are carried as manual input fields until stakeholder sign-off defines the formula.

// FR-3.4: Gratuity only if tenure > 5 years = 50% of Gross Salary × 15 × Tenure ÷ 26
function computeGratuity(grossSalary, tenureYears) {
  const gross = Number(grossSalary || 0);
  const years = Number(tenureYears || 0);
  return years > 5 ? Math.round((0.5 * gross * 15 * years) / 26) : 0;
}

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

  const gratuity = computeGratuity(grossSalary, tenureYears);

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

const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

// Labels the two months a settlement typically spans (LWD's month + the one before),
// matching how HR names the "Salary (August)" / "Salary (July)" style line items.
function settlementMonthLabels(lwd) {
  if (!lwd) return { current: 'Current Month', previous: 'Previous Month' };
  const d = new Date(lwd);
  const current = MONTH_NAMES[d.getMonth()];
  const previous = MONTH_NAMES[(d.getMonth() + 11) % 12];
  return { current, previous };
}

// Sums the detailed settlement breakdown (rec.settlement) — independent of calc()'s
// simplified comp/dues model. This mirrors the HR FnF sheet template line-for-line.
function settlementCalc(rec) {
  const e = (rec.settlement && rec.settlement.earnings) || {};
  const d = (rec.settlement && rec.settlement.deductions) || {};
  const num = (v) => Number(v || 0);

  const earningsTotal = num(e.salaryCurrentMonth) + num(e.salaryPreviousMonth) + num(e.incentives)
    + num(e.taBillPositive) + num(e.gratuity) + num(e.otherReimbursement) + num(e.noticePeriodPay);

  const deductionsTotal = num(d.pfCurrentMonth) + num(d.pfPreviousMonth) + num(d.examDues) + num(d.wfhDues)
    + num(d.nps) + num(d.vpf) + num(d.taBillNegative) + num(d.loanAdvances) + num(d.personalRides)
    + num(d.pluxeeMeal) + num(d.esi) + num(d.tds) + num(d.pt);

  return { earningsTotal, deductionsTotal, netPayable: earningsTotal - deductionsTotal };
}

// Mandatory company holidays — excluded from working/present days only when
// they land on a weekday (a holiday on a weekend is already excluded as a
// weekend, so it must not be double-deducted). [month(0-idx), day].
const MANDATORY_HOLIDAYS = [[0, 1], [0, 26], [7, 15], [9, 2]];

function parseYMD(s) {
  const [y, m, d] = s.split('-').map(Number);
  return { y, m: m - 1, d };
}
function isWeekendUTC(y, m, d) {
  const day = new Date(Date.UTC(y, m, d)).getUTCDay();
  return day === 0 || day === 6;
}
function isMandatoryHoliday(m, d) {
  return MANDATORY_HOLIDAYS.some(([hm, hd]) => hm === m && hd === d);
}
function daysInMonth(y, m) {
  return new Date(Date.UTC(y, m + 1, 0)).getUTCDate();
}
function countEligibleDays(y, m, fromDay, toDay) {
  let count = 0;
  for (let d = fromDay; d <= toDay; d++) {
    if (isWeekendUTC(y, m, d)) continue;
    if (isMandatoryHoliday(m, d)) continue;
    count++;
  }
  return count;
}

// Working Days = every day in the LWD's month except Sat/Sun and the 4
// mandatory holidays (only when they fall on a weekday).
// Present Days = the same count but only through the LWD (inclusive) — days
// after LWD are never counted, even if they'd otherwise be eligible.
// With no LWD yet, falls back to `fallbackDate`'s month, full month for both.
function computeWorkingPresentDays(lwd, fallbackDate) {
  const p = lwd ? parseYMD(lwd) : parseYMD(fallbackDate || new Date().toISOString().slice(0, 10));
  const total = daysInMonth(p.y, p.m);
  const workingDays = countEligibleDays(p.y, p.m, 1, total);
  const presentDays = lwd ? countEligibleDays(p.y, p.m, 1, p.d) : workingDays;
  return { workingDays, presentDays };
}

// Salary Release Condition: when LWD falls in the month after DOR's month, the
// previous month's salary (withheld pending settlement) is released in full
// alongside the current month's salary, pro-rated up to LWD. If DOR and LWD
// fall in the same month, there's nothing withheld — previous month is 0.
function computeSalarySplit(grossSalary, dor, lwd, workingDays, presentDays) {
  const gross = Number(grossSalary || 0);
  const salaryCurrentMonth = workingDays > 0 ? Math.round((gross / workingDays) * presentDays) : 0;
  const spansMonthBoundary = dor && lwd && dor.slice(0, 7) !== lwd.slice(0, 7);
  const salaryPreviousMonth = spansMonthBoundary ? gross : 0;
  return { salaryCurrentMonth, salaryPreviousMonth };
}

// Notice Period Pay = Gross ÷ Working Days × Present Days — same pro-ration
// as the current month's salary, applied to the notice period being paid out.
function computeNoticePay(grossSalary, workingDays, presentDays) {
  const gross = Number(grossSalary || 0);
  return workingDays > 0 ? Math.round((gross / workingDays) * presentDays) : 0;
}

module.exports = { calc, recStatus, recSla, daysBetween, settlementCalc, settlementMonthLabels, computeWorkingPresentDays, computeSalarySplit, computeGratuity, computeNoticePay };
