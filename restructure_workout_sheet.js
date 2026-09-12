/**
 * Workout Tracker Restructure Script (Hip-Safe + 100% Full-Body Micro-Links)
 * 
 * Includes direct hyperlinks to each exercise's row in the "Exercise Guide" tab.
 */

function restructureWorkoutTracker() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var sheet = null;

  // Target main tab (ID 717009429 or first non-copy tab)
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === 717009429) {
      sheet = sheets[i];
      break;
    }
  }

  if (!sheet) {
    for (var j = 0; j < sheets.length; j++) {
      var name = sheets[j].getName().toLowerCase();
      if (!name.includes("copy") && !name.includes("guide")) {
        sheet = sheets[j];
        break;
      }
    }
  }

  if (!sheet) sheet = ss.getActiveSheet();

  sheet.clear();
  sheet.clearFormats();

  // Reset freeze first
  sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0);

  // Look for Exercise Guide to create direct jump links
  var guideSheet = ss.getSheetByName("Exercise Guide");
  var guideGid = guideSheet ? guideSheet.getSheetId() : null;
  var guideRowMap = {};
  if (guideSheet) {
    var gData = guideSheet.getDataRange().getValues();
    for (var g = 1; g < gData.length; g++) {
      var cat = gData[g][0].toString().trim().toLowerCase();
      var gName = gData[g][1].toString().trim().toLowerCase();
      guideRowMap[gName] = g + 1;
      guideRowMap[cat + "_" + gName] = g + 1;
    }
  }

  var exercises = [
    // ==========================================
    // PULL A (Lats, Biceps, Forearm Balance)
    // ==========================================
    { isDivider: true, title: "PULL A WORKOUT (Lats, Biceps & Core)" },
    { day: "Pull A", name: "45° Hyperextension / Back Extension", sets: "3", reps: "10-12", notes: "Glute/hamstring squeeze, zero hip compression" },
    { day: "Pull A", name: "Neutral-Grip Lat Pulldown", sets: "3", reps: "8-12", notes: "Full lat stretch & contraction" },
    { day: "Pull A", name: "Single-Arm Neutral Cable Row", sets: "3", reps: "8-12", notes: "Elbow tucked tight to hip" },
    { day: "Pull A", name: "Cable Face Pulls", sets: "4", reps: "15-20", notes: "External rotation focus, pulls to forehead" },
    { day: "Pull A", name: "Incline Dumbbell Curls", sets: "3", reps: "10-12", notes: "Long head bicep stretch" },
    { day: "Pull A", name: "Dumbbell Hammer Curls", sets: "3", reps: "10-12", notes: "Brachialis focus" },
    { day: "Pull A", name: "Reverse Wrist Extensor Curls", sets: "3", reps: "15-20", notes: "★ MICRO-LINK: Forearm extensors / elbow health" },
    { day: "Core",   name: "Pallof Press", sets: "3", reps: "12-15/side", notes: "Anti-rotation, pelvic stability" },
    { day: "Core",   name: "Deadbugs", sets: "3", reps: "10/side", notes: "Neutral spine & hip flexor control" },

    // ==========================================
    // PUSH A (Chest Focus + Shoulder Health)
    // ==========================================
    { isDivider: true, title: "PUSH A WORKOUT (Chest Focus & Shoulder Health)" },
    { day: "Push A", name: "Shoulder External Rotations", sets: "3", reps: "12-15", notes: "★ MICRO-LINK: Rotator cuff (Infraspinatus/Teres Minor)" },
    { day: "Push A", name: "Flat Bench Press (Barbell or DB)", sets: "4", reps: "6-8", notes: "Heavy main chest compound" },
    { day: "Push A", name: "Standing Overhead Press", sets: "3", reps: "8-10", notes: "Vertical push compound" },
    { day: "Push A", name: "Incline Dumbbell Press", sets: "3", reps: "8-12", notes: "Upper clavicular chest focus" },
    { day: "Push A", name: "Pec Deck or Cable Chest Flyes", sets: "3", reps: "12-15", notes: "Direct chest adduction in stretched position" },
    { day: "Push A", name: "Triceps Rope Pushdowns", sets: "3", reps: "10-12", notes: "Superset with Cable Lateral Raises" },
    { day: "Push A", name: "Cable Lateral Raises", sets: "3", reps: "15-20", notes: "Continuous deltoid tension" },
    { day: "Push A", name: "Push-Up Plus / Serratus Punches", sets: "3", reps: "12-15", notes: "★ MICRO-LINK: Serratus anterior (prevents shoulder impingement)" },
    { day: "Core",   name: "Plank Holds", sets: "3", reps: "45-60s", notes: "Isometric anterior core & TVA" },

    // ==========================================
    // LEGS A (Quad & Glute Drive + Hip Rotators & Shins)
    // ==========================================
    { isDivider: true, title: "LEGS A WORKOUT (Quad & Glute Drive + Hip Rotators)" },
    { day: "Legs A", name: "Heavy Leg Extensions", sets: "3", reps: "10-12", notes: "Main quad lift, 2-sec controlled eccentric" },
    { day: "Legs A", name: "Hip Thrust / Glute Drive Machine", sets: "3", reps: "10-12", notes: "Peak glute loading at 0° hip extension (no socket pinch)" },
    { day: "Legs A", name: "Lying or Standing Leg Curls", sets: "3", reps: "10-12", notes: "Isolated knee flexion hamstring work" },
    { day: "Legs A", name: "Hip Abduction Machine", sets: "3", reps: "12-15", notes: "Gluteus medius & outer hip" },
    { day: "Legs A", name: "Banded Clamshells", sets: "2-3", reps: "15/side", notes: "★ MICRO-LINK: Deep 6 Hip Rotators (centrates ball in socket)" },
    { day: "Legs A", name: "Standing Calf Raises", sets: "4", reps: "10-15", notes: "Gastrocnemius focus (straight knee)" },
    { day: "Legs A", name: "Tibialis Wall Raises", sets: "3", reps: "15-20", notes: "★ MICRO-LINK: Shins / Ankle dorsiflexors (knee & gait health)" },
    { day: "Legs A", name: "Standing Psoas Isometric March", sets: "2", reps: "5x5s/side", notes: "★ MICRO-LINK: Psoas activation above 90° (zero joint shear)" },
    { day: "Core",   name: "Ab Wheel Rollouts", sets: "3", reps: "8-10", notes: "From knees, hips locked neutral at 0°" },

    // ==========================================
    // PULL B (Upper Back, Rear Delts & Carries)
    // ==========================================
    { isDivider: true, title: "PULL B WORKOUT (Upper Back, Rear Delts & Grip)" },
    { day: "Pull B", name: "Chest-Supported T-Bar / Incline DB Row", sets: "4", reps: "8-10", notes: "Replaces BB Row: 0% hip/lower back compressive fatigue" },
    { day: "Pull B", name: "Wide-Grip Lat Pulldown", sets: "3", reps: "8-12", notes: "Upper lat & teres major width" },
    { day: "Pull B", name: "Rear Delt Pec Deck / Face Pulls", sets: "4", reps: "15-20", notes: "Posterior deltoid hypertrophy" },
    { day: "Pull B", name: "Dumbbell or Barbell Shrugs", sets: "3", reps: "10-12", notes: "Upper trapezius overload" },
    { day: "Pull B", name: "Preacher or Machine Curls", sets: "3", reps: "10-12", notes: "Short head bicep peak contraction" },
    { day: "Pull B", name: "Reverse-Grip Cable Curls", sets: "3", reps: "12-15", notes: "Brachioradialis & upper forearm thickness" },
    { day: "Core",   name: "Suitcase Carries (Single-Arm)", sets: "3", reps: "30-40s/side", notes: "Anti-lateral flexion, grip & single-leg hip stability" },

    // ==========================================
    // PUSH B (Shoulders, Triceps & Serratus)
    // ==========================================
    { isDivider: true, title: "PUSH B WORKOUT (Shoulders, Triceps & Serratus)" },
    { day: "Push B", name: "Cable Shoulder External Rotations", sets: "3", reps: "12-15", notes: "★ MICRO-LINK: Rotator cuff stability before overhead work" },
    { day: "Push B", name: "Seated or Standing Overhead Press", sets: "4", reps: "6-8", notes: "Heavy main shoulder compound" },
    { day: "Push B", name: "Incline Barbell or Dumbbell Press", sets: "3", reps: "8-10", notes: "Upper chest angle variation" },
    { day: "Push B", name: "Assisted Dips or Close-Grip Push-ups", sets: "3", reps: "10-12", notes: "Lower chest & triceps compound" },
    { day: "Push B", name: "Dumbbell Lateral Raises", sets: "4", reps: "12-15", notes: "Lateral deltoid cap focus" },
    { day: "Push B", name: "Overhead Cable Triceps Extensions", sets: "3", reps: "10-12", notes: "Long head triceps in full stretch" },
    { day: "Push B", name: "Serratus Cable Punches", sets: "3", reps: "12-15", notes: "★ MICRO-LINK: Serratus anterior protraction" },
    { day: "Core",   name: "Half-Kneeling Pallof Press", sets: "3", reps: "12/side", notes: "Rotational resistance with pelvis locked" },

    // ==========================================
    // LEGS B (Hamstrings, Hip Rehab & Tibialis)
    // ==========================================
    { isDivider: true, title: "LEGS B WORKOUT (Hamstrings, Hip Internal Rotation & Shins)" },
    { day: "Legs B", name: "Dead-Mill Reverse Walking", sets: "3-4", reps: "60-90s", notes: "High VMO quad pump, zero hip compression, synovial fluid flow" },
    { day: "Legs B", name: "Romanian Deadlift (Dumbbells)", sets: "3", reps: "8-10", notes: "Halt at knee level (do NOT touch floor; avoids impingement)" },
    { day: "Legs B", name: "Seated Leg Curls", sets: "3", reps: "10-12", notes: "Hamstring stretch at 90° knee angle" },
    { day: "Legs B", name: "Hip Adduction Machine", sets: "3", reps: "12-15", notes: "Adductor magnus & groin stabilization" },
    { day: "Legs B", name: "Seated Banded Hip Internal Rotation", sets: "2-3", reps: "12-15", notes: "★ MICRO-LINK: Gluteus minimus (restores impingement-free socket glide)" },
    { day: "Legs B", name: "Seated Calf Raises", sets: "4", reps: "12-15", notes: "Soleus focus (bent knee)" },
    { day: "Legs B", name: "Tibialis Wall Raises", sets: "3", reps: "15-20", notes: "★ MICRO-LINK: Tibialis anterior (balances calf raises)" },
    { day: "Core",   name: "Side Planks", sets: "3", reps: "30-45s/side", notes: "Quadratus lumborum & lateral hip chain" }
  ];

  var totalCols = 5 + (4 * 12);
  var row1 = ["EXERCISE", "", "DETAILS & TARGETS", "", ""];
  var row2 = ["Day", "Exercise", "Sets", "Target Reps", "Notes / Focus"];

  for (var w = 1; w <= 4; w++) {
    row1.push("WEEK " + w + " — BIKASH", "", "", "", "", "");
    row2.push("Set 1 (lbs)", "Reps", "Set 2 (lbs)", "Reps", "Set 3 (lbs)", "Reps");

    row1.push("WEEK " + w + " — LIZ", "", "", "", "", "");
    row2.push("Set 1 (lbs)", "Reps", "Set 2 (lbs)", "Reps", "Set 3 (lbs)", "Reps");
  }

  var allData = [row1, row2];

  for (var k = 0; k < exercises.length; k++) {
    var ex = exercises[k];
    if (ex.isDivider) {
      var divRow = ["", "▶ " + ex.title, "", "", ""];
      for (var col = 5; col < totalCols; col++) divRow.push("");
      allData.push(divRow);
    } else {
      var displayName = ex.name;
      if (guideGid) {
        var cleanCat = ex.day.toString().trim().toLowerCase();
        var cleanName = ex.name.toString().trim().toLowerCase();
        var targetRow = guideRowMap[cleanCat + "_" + cleanName] || guideRowMap[cleanName];
        if (targetRow) {
          displayName = '=HYPERLINK("#gid=' + guideGid + '&range=B' + targetRow + '", "' + ex.name + '")';
        }
      }
      var row = [ex.day, displayName, ex.sets, ex.reps, ex.notes];
      for (var rest = 0; rest < 4 * 12; rest++) row.push("");
      allData.push(row);
    }
  }

  // Set values & formulas
  sheet.getRange(1, 1, allData.length, totalCols).setValues(allData);

  // Freeze top 2 rows & first 2 columns (done BEFORE merges)
  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(2);

  // Row 1 Merges (never crosses column 2 boundary)
  sheet.getRange(1, 1, 1, 2).merge(); // A1:B1 (inside frozen cols)
  sheet.getRange(1, 3, 1, 3).merge(); // C1:E1 (outside frozen cols)

  // Merges for Week blocks (Cols 6+)
  var colIdx = 6;
  for (var w = 1; w <= 4; w++) {
    sheet.getRange(1, colIdx, 1, 6).merge(); // Bikash Week w
    colIdx += 6;
    sheet.getRange(1, colIdx, 1, 6).merge(); // Liz Week w
    colIdx += 6;
  }

  // Row 1 Styling
  sheet.getRange(1, 1, 1, 2).setBackground("#0f172a").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center").setFontSize(11);
  sheet.getRange(1, 3, 1, 3).setBackground("#0f172a").setFontColor("#94a3b8").setFontWeight("bold").setHorizontalAlignment("center").setFontSize(10);
  for (var w = 0; w < 4; w++) {
    var bCol = 6 + (w * 12);
    var lCol = bCol + 6;
    sheet.getRange(1, bCol, 1, 6).setBackground("#1e40af").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center").setFontSize(11);
    sheet.getRange(1, lCol, 1, 6).setBackground("#6b21a8").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center").setFontSize(11);
  }

  // Row 2 Styling
  sheet.getRange(2, 1, 1, 5).setBackground("#334155").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center").setFontSize(10);
  for (var w = 0; w < 4; w++) {
    var bCol = 6 + (w * 12);
    var lCol = bCol + 6;
    sheet.getRange(2, bCol, 1, 6).setBackground("#2563eb").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center").setFontSize(9);
    sheet.getRange(2, lCol, 1, 6).setBackground("#7c3aed").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center").setFontSize(9);
  }

  // Data formatting
  for (var r = 0; r < exercises.length; r++) {
    var rowNum = r + 3;
    var ex = exercises[r];
    if (ex.isDivider) {
      sheet.getRange(rowNum, 1, 1, totalCols).setBackground("#0f172a").setFontColor("#38bdf8").setFontWeight("bold").setFontSize(10);
      sheet.getRange(rowNum, 2).setHorizontalAlignment("left");
    } else {
      var isMicroLink = ex.notes && ex.notes.includes("MICRO-LINK");
      sheet.getRange(rowNum, 1, 1, totalCols).setBackground(isMicroLink ? "#f0fdf4" : (r % 2 === 0 ? "#ffffff" : "#f8fafc"));
      sheet.getRange(rowNum, 1).setHorizontalAlignment("center").setFontWeight("bold");
      sheet.getRange(rowNum, 2).setHorizontalAlignment("left").setFontWeight("bold");
      
      if (guideGid) {
        sheet.getRange(rowNum, 2).setFontLine("underline");
      }

      if (isMicroLink) {
        sheet.getRange(rowNum, 2).setFontColor("#15803d");
      } else if (guideGid) {
        sheet.getRange(rowNum, 2).setFontColor("#1d4ed8");
      }

      sheet.getRange(rowNum, 3, 1, 2).setHorizontalAlignment("center");
      sheet.getRange(rowNum, 5).setHorizontalAlignment("left").setFontColor(isMicroLink ? "#166534" : "#64748b");
      sheet.getRange(rowNum, 6, 1, totalCols - 5).setHorizontalAlignment("center");
    }
  }

  // Column Widths
  sheet.setColumnWidth(1, 80);   // Day
  sheet.setColumnWidth(2, 260);  // Exercise
  sheet.setColumnWidth(3, 75);   // Sets
  sheet.setColumnWidth(4, 100);  // Target Reps
  sheet.setColumnWidth(5, 330);  // Notes / Focus
  for (var c = 6; c <= totalCols; c++) {
    sheet.setColumnWidth(c, 70); // Set 1-3 Wt and Reps
  }

  // Grid borders
  sheet.getRange(1, 1, allData.length, totalCols).setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
}
