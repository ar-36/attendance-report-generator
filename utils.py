from datetime import datetime

def get_settings():
  return {
    'margin_in_mins': 10,
    'colors': {
      'LATE': { 'text':'92400e', 'bg':'fef3c7' },
      'PUBLIC_HOLIDAY': { 'text':'730e90', 'bg':'f3d2fd' },
      'PRESENT': { 'text':'0b8721', 'bg':'c7fed1' },
      'WFH': { 'text':'0c2a91', 'bg':'c5ddfd' },
      'LEAVE_NOT_APPLIED': { 'text':'991b1b', 'bg':'fecaca' },
      'ALL_LEAVES': { 'text':'0b8689', 'bg':'c7fdfe' }
    },
  }


def generate_filename(emp_name, start_date, end_date):
  start_dt = datetime.strptime(start_date, '%Y-%m-%d')
  end_dt = datetime.strptime(end_date, '%Y-%m-%d')

  start_str = start_dt.strftime('%b%d')
  end_str = end_dt.strftime('%b%d')
  year = end_dt.strftime('%Y')

  filename = f"Attendance_{emp_name}_{start_str}-{end_str}_{year}.xlsx"
  return filename


def minutes_to_readable(value):
  if value < 60:
    return f'{int(value)} m'
  else:
    hours = value / 60
    if hours.is_integer():
      return f'{int(hours)} h'
    else:
      return f'{hours:.1f} h'

def get_month_year(datetime_str):
  dt = datetime.strptime(datetime_str, '%Y-%m-%d')
  return dt.strftime('%B %Y')


def get_attendance_report_query(emp_id, start_date, end_date, margin_in_mins):
  return f"""
    SELECT 
      e.emp_other_id AS employee_id,
      CONCAT_WS(' ', NULLIF(MIN(e.emp_firstname), ''), NULLIF(MIN(e.emp_middle_name), ''), NULLIF(MIN(e.emp_lastname), '')) AS employee_name,
      DATE_FORMAT(DATE_ADD(GREATEST('{start_date}', e.joined_date), INTERVAL n.n DAY), '%Y-%m-%d') AS attendance_date,
      sh.shift_interval AS shift,
      DATE_FORMAT(MIN(a.entry), '%H:%i:%s') AS check_in_time,

      CASE 
          WHEN MIN(a.entry) IS NOT NULL THEN
              GREATEST(
                  TIMESTAMPDIFF(
                      MINUTE,
                      STR_TO_DATE(CONCAT(DATE_ADD(GREATEST('{start_date}', e.joined_date), INTERVAL n.n DAY), ' ',
                          CONVERT(SUBSTRING_INDEX(sh.shift_interval, '-', 1) USING utf8mb4) COLLATE utf8mb4_unicode_ci), '%Y-%m-%d %l%p'),
                      MIN(a.entry)
                  ) - {margin_in_mins},
                  0
              )
          ELSE NULL
      END AS late_in_mins,

      CASE
          WHEN ph.id IS NOT NULL THEN 'PUBLIC_HOLIDAY'
          WHEN COALESCE(MAX(l.half_day_leave),0) = 1 AND COALESCE(MAX(l.half_day_wfh),0) = 1 THEN 'HALF_DAY_LEAVE_AND_WFH'
            WHEN COALESCE(MAX(l.half_day_leave),0) = 1 AND MIN(a.entry) IS NOT NULL THEN 'HALF_DAY_LEAVE'
          WHEN COALESCE(MAX(l.half_day_wfh),0) = 1 AND MIN(a.entry) IS NOT NULL THEN 'HALF_DAY_WFH'
          WHEN MIN(a.entry) IS NOT NULL AND
            GREATEST(
                TIMESTAMPDIFF(
                    MINUTE,
                    STR_TO_DATE(CONCAT(DATE_ADD(GREATEST('{start_date}', e.joined_date), INTERVAL n.n DAY), ' ',
                        CONVERT(SUBSTRING_INDEX(sh.shift_interval, '-', 1) USING utf8mb4) COLLATE utf8mb4_unicode_ci), '%Y-%m-%d %l%p'),
                    MIN(a.entry)
                ) - {margin_in_mins},
                0
            ) > {margin_in_mins} THEN 'LATE'
          WHEN MIN(a.entry) IS NOT NULL THEN 'PRESENT'
          WHEN FIND_IN_SET('1', COALESCE(MAX(l.leave_types), '')) > 0 THEN 'ANNUAL_LEAVE'
          WHEN FIND_IN_SET('2', COALESCE(MAX(l.leave_types), '')) > 0 THEN 'SICK_LEAVE'
          WHEN FIND_IN_SET('4', COALESCE(MAX(l.leave_types), '')) > 0 THEN 'WFH'
          ELSE 'LEAVE_NOT_APPLIED'
      END AS status

    FROM (
      SELECT a.N + b.N * 10 + c.N * 100 AS n
      FROM (SELECT 0 N UNION ALL SELECT 1 UNION ALL SELECT 2 UNION ALL SELECT 3 UNION ALL SELECT 4 
            UNION ALL SELECT 5 UNION ALL SELECT 6 UNION ALL SELECT 7 UNION ALL SELECT 8 UNION ALL SELECT 9) a
      CROSS JOIN (SELECT 0 N UNION ALL SELECT 1 UNION ALL SELECT 2 UNION ALL SELECT 3 UNION ALL SELECT 4 
                  UNION ALL SELECT 5 UNION ALL SELECT 6 UNION ALL SELECT 7 UNION ALL SELECT 8 UNION ALL SELECT 9) b
      CROSS JOIN (SELECT 0 N UNION ALL SELECT 1 UNION ALL SELECT 2 UNION ALL SELECT 3 UNION ALL SELECT 4 
                  UNION ALL SELECT 5 UNION ALL SELECT 6 UNION ALL SELECT 7 UNION ALL SELECT 8 UNION ALL SELECT 9) c
    ) n
    JOIN hs_hr_employee e 
      ON e.emp_status IN (2,3)
      AND e.emp_other_id IS NOT NULL
      AND e.emp_other_id = {emp_id}


    LEFT JOIN employees_shift_history sh 
      ON sh.emp_other_id = e.emp_other_id
      AND sh.shift_changed_date = (
          SELECT MAX(s2.shift_changed_date)
          FROM employees_shift_history s2
          WHERE s2.emp_other_id = e.emp_other_id
            AND s2.shift_changed_date <= DATE_ADD(GREATEST('{start_date}', e.joined_date), INTERVAL n.n DAY)
      )

    LEFT JOIN employee_attendance a 
      ON a.userId = e.emp_other_id
      AND (
        (
          -- Normal same-day shift
          DATE(a.entry) = DATE_ADD(GREATEST('{start_date}', e.joined_date), INTERVAL n.n DAY)
          AND a.entry BETWEEN
              DATE_SUB(
                TIMESTAMP(
                  DATE_ADD(GREATEST('{start_date}', e.joined_date), INTERVAL n.n DAY),
                  STR_TO_DATE(CONVERT(SUBSTRING_INDEX(sh.shift_interval,'-',1) USING utf8mb4) COLLATE utf8mb4_unicode_ci, '%l%p')
                ), INTERVAL 2 HOUR
              )
              AND
              DATE_ADD(
                TIMESTAMP(
                  DATE_ADD(GREATEST('{start_date}', e.joined_date), INTERVAL n.n DAY),
                  STR_TO_DATE(CONVERT(SUBSTRING_INDEX(sh.shift_interval,'-',1) USING utf8mb4) COLLATE utf8mb4_unicode_ci, '%l%p')
                ), INTERVAL 12 HOUR
              )
        )
        OR
        (
          -- Cross-midnight shift (e.g. 9PM-6AM)
          STR_TO_DATE(SUBSTRING_INDEX(sh.shift_interval,'-',-1), '%l%p')
          <= STR_TO_DATE(SUBSTRING_INDEX(sh.shift_interval,'-',1), '%l%p')
          AND DATE(a.entry) = DATE_ADD(GREATEST('{start_date}', e.joined_date), INTERVAL (n.n + 1) DAY)
          AND a.entry BETWEEN
              DATE_SUB(
                TIMESTAMP(
                  DATE_ADD(GREATEST('{start_date}', e.joined_date), INTERVAL n.n DAY),
                  STR_TO_DATE(CONVERT(SUBSTRING_INDEX(sh.shift_interval,'-',1) USING utf8mb4) COLLATE utf8mb4_unicode_ci, '%l%p')
                ), INTERVAL 2 HOUR
              )
              AND
              DATE_ADD(
                TIMESTAMP(
                  DATE_ADD(GREATEST('{start_date}', e.joined_date), INTERVAL (n.n + 1) DAY),
                  STR_TO_DATE(CONVERT(SUBSTRING_INDEX(sh.shift_interval,'-',-1) USING utf8mb4) COLLATE utf8mb4_unicode_ci, '%l%p')
                ), INTERVAL 2 HOUR
              )
        )
    )

    LEFT JOIN (
      SELECT emp_number, date,
              GROUP_CONCAT(leave_type_id ORDER BY leave_type_id) AS leave_types,
              SUM(CASE WHEN leave_type_id = 1 AND length_days < 1 THEN 1 ELSE 0 END) AS half_day_leave,
              SUM(CASE WHEN leave_type_id = 4 AND length_days < 1 THEN 1 ELSE 0 END) AS half_day_wfh
      FROM ohrm_leave
      WHERE status IN (1,2,3)
      GROUP BY emp_number, date
    ) l
    ON l.emp_number = e.emp_number
    AND (
          l.date = DATE_ADD(GREATEST('{start_date}', e.joined_date), INTERVAL n.n DAY)
          OR
          (
            STR_TO_DATE(
              CONVERT(SUBSTRING_INDEX(sh.shift_interval,'-',-1) USING utf8mb4) COLLATE utf8mb4_unicode_ci,
              '%l%p'
            )
            <= STR_TO_DATE(
              CONVERT(SUBSTRING_INDEX(sh.shift_interval,'-',1) USING utf8mb4) COLLATE utf8mb4_unicode_ci,
              '%l%p'
            )
            AND l.date = DATE_ADD(GREATEST('{start_date}', e.joined_date), INTERVAL (n.n + 1) DAY)
          )
      )

    LEFT JOIN attendance_public_holidays ph
      ON ph.holiday_date = DATE_ADD(GREATEST('{start_date}', e.joined_date), INTERVAL n.n DAY)
      # AND ph.shift_time = sh.shift_interval

    WHERE DATE_ADD(GREATEST('{start_date}', e.joined_date), INTERVAL n.n DAY) <= '{end_date}'
    AND WEEKDAY(DATE_ADD(GREATEST('{start_date}', e.joined_date), INTERVAL n.n DAY)) BETWEEN 0 AND 4

    GROUP BY e.emp_other_id, n.n, sh.shift_interval, ph.id, e.joined_date
    ORDER BY attendance_date;
  """