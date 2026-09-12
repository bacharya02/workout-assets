/**
 * Visual Exercise Guide Creator & Auto-Hyperlinker
 * 
 * Includes:
 * 1. createExerciseGuideTab() - Builds "Exercise Guide" with verified images, videos & cues.
 * 2. addExerciseHyperlinks()   - Adds clickable links to Column B in your Workout Tracker
 *                               to jump directly to each exercise's row in the guide.
 */

function createExerciseGuideTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "Exercise Guide";
  var sheet = ss.getSheetByName(sheetName);

  if (sheet) {
    sheet.clear();
    sheet.clearFormats();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0);

  var cdn = "https://cdn.jsdelivr.net/gh/bacharya02/workout-assets@main/images/";

  // Array of: [Category, Name, Muscles, Image File, Video Search Query, Form Cues, Hip Safety Rationale]
  var data = [
    // PULL A
    ["Pull A", "45° Hyperextension / Back Extension", "Glutes, Hamstrings, Lower Back", "back-extension-peak.webp", "how+to+do+45+degree+back+extension+for+glutes", "Set pad right below hip crease. Round upper back slightly to focus on glutes/hamstrings. Hinge down at hips, squeeze glutes hard to lift torso to straight line. Do NOT hyperextend lower back.", "Zero compressive axial loading on the hip socket; trains posterior chain safely."],
    ["Pull A", "Neutral-Grip Lat Pulldown", "Lats (Latissimus Dorsi), Biceps", "close-grip-lat-pulldown-peak.webp", "how+to+do+neutral+grip+lat+pulldown", "Use parallel/neutral handles (palms facing each other). Lean back ~10-15 degrees. Pull elbows straight down toward your hips, bringing bar to collarbone. Squeeze lats for 1s, control return.", "Neutral grip spares shoulders and locks tension on lats without twisting joints."],
    ["Pull A", "Single-Arm Neutral Cable Row", "Lower & Mid Lats, Rhomboids", "bent-over-db-row-peak.webp", "how+to+do+single+arm+cable+row+lat+bias", "Set cable at belly-button height. Sit tall. Reach forward to stretch the lat, then drive elbow back skimming your ribs. Keep chest proud—do not rotate torso excessively.", "Single-arm allows full lat stretch and contraction without spinal compression."],
    ["Pull A", "Cable Face Pulls", "Rear Delts, Rotator Cuff, Upper Traps", "dumbbell-face-pull-peak.webp", "how+to+do+face+pulls+proper+form", "Set rope attachment at eye height. Grip with thumbs backward. Pull hands toward your ears/forehead while pulling rope ends apart. Rotate wrists back so knuckles face behind you.", "Essential posture builder and rotator cuff external rotation."],
    ["Pull A", "Incline Dumbbell Curls", "Biceps (Long Head)", "incline-db-curl-peak.webp", "how+to+do+incline+dumbbell+curls", "Set bench to 60-degree angle. Lean back and let arms hang straight down behind torso. Curl up while keeping elbows pinned in place. Full stretch at the bottom.", "Places biceps in a full deep stretch for superior hypertrophy."],
    ["Pull A", "Dumbbell Hammer Curls", "Brachialis, Brachioradialis", "hammer-curl-peak.webp", "how+to+do+hammer+curls", "Hold dumbbells with neutral grip (palms facing each other, like holding a hammer). Curl upward without swinging torso. Lower slowly.", "Builds forearm thickness and pushes the bicep peak upward."],
    ["Pull A", "Reverse Wrist Extensor Curls", "★ Forearm Extensors", "barbell-wrist-curl-peak.webp", "how+to+do+reverse+wrist+curls", "Rest forearms flat on bench or thighs, holding light dumbbells with palms facing DOWN. Let hands hang off edge, then extend wrists upward toward ceiling. Slow controlled reps.", "★ MICRO-LINK: Balances heavy pulling grip and prevents tennis elbow (lateral epicondylitis)."],
    ["Pull A", "Pallof Press", "Core (Transverse Abdominis, Obliques)", "cable-pallof-press-peak.webp", "how+to+do+pallof+press", "Stand sideways to cable machine. Hold handle at center of chest with both hands. Press straight out away from body; resist the cable's attempt to rotate you. Hold 2s, bring back.", "Anti-rotation core power with 0% rotational torque on the hip socket."],
    ["Pull A", "Deadbugs", "Deep Core, Anti-Extension", "dead-bug-hold-main.webp", "how+to+do+deadbug+exercise", "Lie flat on back. Push lower back firmly into the floor (no gap!). Arms straight up, knees at 90°. Slowly extend right arm overhead and left leg straight out. Return and alternate.", "Strengthens abdominal wall with zero hip flexor yanking or spinal shear."],

    // PUSH A
    ["Push A", "Shoulder External Rotations", "★ Rotator Cuff (Infraspinatus, Teres Minor)", "cable-external-rotation-peak.webp", "cable+shoulder+external+rotation+elbow+at+side", "Stand sideways to cable at elbow height (or lie on side with light DB). Keep elbow pinned tightly to your side at 90°. Rotate forearm outward away from stomach. Slow and smooth.", "★ MICRO-LINK: Strengthens the small stabilizers that prevent rotator cuff tears."],
    ["Push A", "Flat Bench Press (Barbell or DB)", "Chest, Front Delts, Triceps", "bench-press-peak.webp", "how+to+flat+bench+press+proper+form", "Lie on bench with eyes under bar. Grip slightly outside shoulder width. Squeeze shoulder blades together and down. Lower bar to mid-chest, press up in a slight backward arc.", "Primary upper body horizontal pushing compound."],
    ["Push A", "Standing Overhead Press", "Shoulders, Triceps, Upper Chest", "behind-the-neck-press-peak.webp", "how+to+overhead+press+proper+form", "Grip bar just outside shoulders. Brace core and squeeze glutes tight. Press bar straight up overhead, moving chin back as bar passes, then locking out with bar over mid-foot.", "Pure vertical pushing strength and core stability."],
    ["Push A", "Incline Dumbbell Press", "Upper Chest (Clavicular Head), Front Delts", "incline-bench-press-peak.webp", "how+to+incline+dumbbell+press", "Set bench to 30-45 degrees. Press dumbbells up over upper chest with palms angled at ~45 degrees. Lower with control until dumbbells touch outside of chest.", "Targets upper chest fibers to fill out the collarbone region."],
    ["Push A", "Pec Deck or Cable Chest Flyes", "Chest (Inner/Sternal Head Isolation)", "cable-fly-peak.webp", "how+to+use+pec+deck+machine+chest+fly", "Adjust seat so handles are at mid-chest height. Keep slight bend in elbows. Bring hands together in front like hugging a barrel. Squeeze chest hard for 1 second.", "Direct chest adduction in fully stretched and contracted positions."],
    ["Push A", "Triceps Rope Pushdowns", "Triceps (Lateral & Medial Heads)", "banded-triceps-stretch-main.webp", "how+to+tricep+rope+pushdown", "Pin elbows to ribcage. Push rope downward and flare the ends outward at the bottom. Squeeze triceps, then return to 90 degrees with control.", "Isolates the triceps without stressing wrists or elbows."],
    ["Push A", "Cable Lateral Raises", "Side Deltoids (Lateral Cap)", "cable-lateral-raise-peak.webp", "how+to+cable+lateral-raise", "Set cable at lowest setting. Stand tall, raise cable handle out to side up to shoulder level. Lead slightly with elbow. Cable provides constant tension unlike dumbbells.", "Creates round shoulder width with continuous resistance curve."],
    ["Push A", "Push-Up Plus / Serratus Punches", "★ Serratus Anterior", "archer-push-ups-peak.webp", "how+to+do+push+up+plus+for+serratus", "In a push-up or high plank position (arms straight). Without bending your elbows, push your upper back toward the ceiling, spreading shoulder blades wide apart. Hold 2s, relax.", "★ MICRO-LINK: Prevents shoulder impingement and scapular winging."],
    ["Push A", "Plank Holds", "Anterior Core & Transverse Abdominis", "plank-main.webp", "how+to+proper+plank+form", "Rest on forearms and toes. Tuck pelvis slightly (posterior tilt), squeeze glutes, brace abs as if someone is about to punch your stomach. Breathe through your nose.", "Strengthens the abdominal brace without flexing or pinching the hips."],

    // LEGS A
    ["Legs A", "Heavy Leg Extensions", "Quadriceps (All 4 heads)", "leg-extension-peak.webp", "how+to+do+leg+extensions+proper+form", "Align knee joint with machine pivot point. Back flat against pad. Extend legs smoothly to lockout, pause 1 second at top, lower under control for 2 full seconds.", "Pure quad hypertrophy without ANY compressive load on the hip joint."],
    ["Legs A", "Hip Thrust / Glute Drive Machine", "Gluteus Maximus", "hip-thrust-peak.webp", "how+to+do+hip+thrust+machine", "Place pad/belt over hips, upper back on pad. Feet shoulder-width. Drive through heels to full hip extension (flat 0° angle). Squeeze glutes at top. Chin tucked.", "Highest glute activation possible; zero socket impingement because load peaks at extension."],
    ["Legs A", "Lying or Standing Leg Curls", "Hamstrings (Knee Flexion)", "ball-leg-curl-peak.webp", "how+to+do+lying+leg+curls", "Pad rests on lower calves just above heels. Keep hips pressed flat into bench (don't arch back). Curl heels toward glutes, pause, and lower slowly.", "Direct hamstring knee flexion isolation with zero hip compression."],
    ["Legs A", "Hip Abduction Machine", "Gluteus Medius & Outer Hip", "banded-seated-hip-abduction-peak.webp", "how+to+use+hip+abduction+machine", "Sit upright with pads on outside of knees. Push knees outward against resistance as far as possible. Hold for 2 seconds at peak spread, then slow return.", "Strengthens lateral pelvic stabilizers."],
    ["Legs A", "Banded Clamshells", "★ Deep 6 Hip Rotators (Hip Rotator Cuff)", "banded-clamshell-peak.webp", "how+to+do+clamshells+exercise+glutes", "Lie on side with loop band above knees. Knees bent 45°, feet glued together. Keeping feet touching, rotate top knee open toward ceiling without letting your pelvis roll backward. Hold 2s.", "★ MICRO-LINK: Fires the gemelli, obturators, and piriformis to keep the ball centered in the socket."],
    ["Legs A", "Standing Calf Raises", "Gastrocnemius (Upper Calf)", "barbell-calf-raise-peak.webp", "how+to+standing+calf+raises", "Knees kept completely straight. Balls of feet on edge of block. Drop heels for full stretch, then press up high onto big toes. Pause 2s at the top.", "Straight knee targets the large gastrocnemius calf muscle."],
    ["Legs A", "Tibialis Wall Raises", "★ Tibialis Anterior (Shins)", "bodyweight-calf-raise-peak.webp", "how+to+do+tibialis+raises+against+wall", "Lean buttocks flat against a wall, feet ~12-18 inches out in front with legs straight. Flex your toes upward toward your shins as high as possible. Pause, lower, repeat.", "★ MICRO-LINK: Balances calves, absorbs walking impact, and stabilizes knees & hips."],
    ["Legs A", "Standing Psoas Isometric March", "★ Psoas Major (Deep Hip Flexor)", "high-knees-main.webp", "psoas+march+isometric+hold", "Stand tall near wall. Lift knee up until thigh is above 90 degrees (above parallel). Place hand on knee and press down firmly while pushing up with knee for 5s hold. 0 movement.", "★ MICRO-LINK: Activates deep psoas in its pure functional zone above 90° with ZERO joint shear."],
    ["Legs A", "Ab Wheel Rollouts", "Core (Rectus Abdominis & Lats)", "ab-wheel-rollout-peak.webp", "how+to+do+ab+wheel+rollout+proper+form", "Kneel on pad. Squeeze glutes and round upper back slightly (hollow body). Roll wheel forward slowly only as far as you can maintain a flat back. Pull back using abs, not hips.", "Top tier anti-extension core strength without hip flexor yanking."],

    // PULL B
    ["Pull B", "Chest-Supported T-Bar / Incline DB Row", "Upper Back, Rhomboids, Lats", "barbell-row-peak.webp", "how+to+chest+supported+t+bar+row", "Lie chest-down on incline bench or machine pad. Grab handles with overhand or neutral grip. Pull elbows backward past your torso, squeezing shoulder blades together.", "Completely eliminates lower back and hip socket fatigue while blasting the upper back."],
    ["Pull B", "Wide-Grip Lat Pulldown", "Upper Lats, Teres Major", "lat-pulldown-peak.webp", "how+to+wide+grip+lat+pulldown", "Grip bar ~1.5x shoulder width. Sit under thigh pads. Pull bar down to upper chest, leading with your elbows. Focus on spreading lats wide at the top.", "Builds upper back V-taper width."],
    ["Pull B", "Rear Delt Pec Deck / Face Pulls", "Posterior Deltoids", "dumbbell-reverse-fly-peak.webp", "how+to+reverse+pec+deck+rear+delt+fly", "Sit facing machine pad. Handles at shoulder height. Keeping elbows slightly bent, fly arms backward in a wide horizontal arc. Squeeze rear delts at peak.", "Fixes rounded forward shoulders and bulletproofs the upper back."],
    ["Pull B", "Dumbbell or Barbell Shrugs", "Upper Trapezius", "behind-the-back-barbell-shrug-peak.webp", "how+to+do+dumbbell+shrugs+proper+form", "Hold dumbbells at your sides. Elevate shoulders straight up toward your ears. Hold for 1 full second at top. Lower slowly. Do NOT roll shoulders backward or forward.", "Builds neck and upper trapezius thickness."],
    ["Pull B", "Preacher or Machine Curls", "Biceps (Short Head / Peak)", "barbell-preacher-curl-peak.webp", "how+to+preacher+curl+proper+form", "Armpits snug against top of pad. Keep upper arms completely glued to pad. Curl bar up to chin height, contract bicep hard, and lower with control.", "Eliminates all shoulder momentum and isolates the bicep completely."],
    ["Pull B", "Reverse-Grip Cable Curls", "Brachioradialis & Forearms", "db-reverse-curl-peak.webp", "how+to+reverse+grip+cable+curl", "Attach straight or EZ bar to low cable. Grip overhand (palms facing down). Curl bar upward toward collarbones. Keep wrists rigid and straight throughout.", "Builds dense outer forearms and strengthens elbow joints."],
    ["Pull B", "Suitcase Carries (Single-Arm)", "Obliques, QL, Grip, Pelvic Stabilizers", "dumbbell-farmers-walk-main.webp", "how+to+do+suitcase+carries", "Hold a heavy dumbbell in ONE hand only like a suitcase. Stand tall with shoulders level. Walk forward slowly; resist leaning or tilting toward the weight.", "Elite anti-lateral flexion core and single-leg hip stability exercise."],

    // PUSH B
    ["Push B", "Cable Shoulder External Rotations", "★ Shoulder Rotator Cuff", "cable-external-rotation-peak.webp", "cable+shoulder+external+rotation+elbow+at+side", "Elbow tucked into side at 90 degrees with a towel under arm. Rotate hand outward away from body. Smooth, slow reps.", "★ MICRO-LINK: Warms up and protects rotator cuff before heavy overhead pressing."],
    ["Push B", "Seated or Standing Overhead Press", "Shoulders, Triceps", "bodyweight-overhead-press-peak.webp", "how+to+seated+dumbbell+shoulder+press", "Press dumbbells or barbell straight up until arms are fully extended overhead. Lower dumbbells to ear level under control.", "Primary vertical push for full shoulder cap development."],
    ["Push B", "Incline Barbell or Dumbbell Press", "Upper Pectorals", "incline-bench-press-peak.webp", "how+to+incline+bench+press", "Bench at 30 degrees. Lower bar to upper chest smoothly, press up locking out with chest flexed.", "Upper chest focus with barbell or dumbbell variety."],
    ["Push B", "Assisted Dips or Close-Grip Push-ups", "Lower Chest, Triceps", "assisted-dips-peak.webp", "how+to+do+dips+proper+form", "Lean torso forward 15-20 degrees for chest emphasis. Lower until elbows hit 90 degrees. Press through palms back to top lockout.", "Excellent compound builder for lower chest and triceps."],
    ["Push B", "Dumbbell Lateral Raises", "Side Deltoids", "seated-dumbbell-lateral-raise-peak.webp", "how+to+do+dumbbell+lateral+raise", "Slight forward torso hinge. Raise dumbbells out to sides leading with your elbows. Imagine pouring water out of a pitcher at the top. Pause and lower.", "Standard lateral head shoulder builder."],
    ["Push B", "Overhead Cable Triceps Extensions", "Triceps (Long Head)", "kettlebell-overhead-tricep-extension-peak.webp", "how+to+overhead+cable+tricep+extension", "Set rope high or mid. Step forward facing away from machine. Elbows up beside head. Extend forearms straight forward, squeezing triceps at full reach.", "Full stretch on the long head of the triceps (largest tricep muscle)."],
    ["Push B", "Serratus Cable Punches", "★ Serratus Anterior", "archer-push-ups-peak.webp", "how+to+do+serratus+punches", "Hold cable handle at chest height facing away. Punch straight forward until arm is extended, then push forward an additional 2 inches through your shoulder blade.", "★ MICRO-LINK: Directly activates serratus to glide the scapula forward."],
    ["Push B", "Half-Kneeling Pallof Press", "Core (Anti-Rotation)", "cable-pallof-press-peak.webp", "half+kneeling+pallof+press", "Drop inside knee onto cushion (outside leg up at 90 degrees). Hold cable at chest, press straight out and hold 2s against the rotational pull. Return to chest.", "Locks the pelvis completely in place so the hip joint cannot rotate or twist."],

    // LEGS B
    ["Legs B", "Dead-Mill Reverse Walking", "Quads (VMO), Patellar Tendon, Synovial Fluid", "incline-treadmill-walk-main.webp", "backwards+treadmill+walking+kneesovertoes", "Leave treadmill completely turned OFF. Hold handrails for balance. Walk backwards by pushing the belt with your feet (toe-to-heel strike, driving quads).", "Unmatched knee and hip rehab: forces blood & synovial fluid into joints with 0% joint compression."],
    ["Legs B", "Romanian Deadlift (Dumbbells)", "Hamstrings, Glutes", "banded-romanian-deadlift-peak.webp", "how+to+do+dumbbell+romanian+deadlift", "Hold dumbbells against front of thighs. Soft knee bend. Push hips back as if touching a wall behind you. STOP at knee level (do NOT go to floor). Drive hips forward to stand.", "Halt at knee level preserves hamstring training while avoiding deep hip impingement."],
    ["Legs B", "Seated Leg Curls", "Hamstrings (Lengthened Position)", "seated-leg-curl-peak.webp", "how+to+seated+leg+curl", "Thigh pad locked down firmly over quads. Curl heels down and under seat as far as possible. Hold 1s, allow slow 2s return.", "Studies show seated leg curls produce greater hamstring hypertrophy than lying leg curls."],
    ["Legs B", "Hip Adduction Machine", "Adductors (Inner Thigh / Groin)", "banded-standing-hip-adduction-peak.webp", "how+to+use+hip+adduction+machine", "Pads on inside of knees. Spread legs wide, then squeeze knees together toward centerline under control. Hold 1s, slow controlled return.", "Strengthens adductor magnus and stabilizes pelvis against lateral shift."],
    ["Legs B", "Seated Banded Hip Internal Rotation", "★ Gluteus Minimus & Anterior Medius", "banded-clamshell-start.webp", "seated+banded+hip+internal+rotation", "Sit on bench with resistance band looped around both ankles. Place yoga block or foam roller between knees. Keeping knees clamped on block, kick both ankles outward away from each other.", "★ MICRO-LINK: Restores active internal rotation so the ball glides smoothly in the socket without pinching."],
    ["Legs B", "Seated Calf Raises", "Soleus (Deep Lower Calf)", "seated-calf-raise-peak.webp", "how+to+seated+calf+raise", "Knees bent at 90 degrees with pad over lower thighs. Drop heels deep for a full stretch, press through balls of feet to full height.", "Bent knee isolates the soleus muscle (which straight-leg calf raises cannot fully target)."],
    ["Legs B", "Tibialis Wall Raises", "★ Tibialis Anterior (Shins)", "bodyweight-calf-raise-peak.webp", "how+to+do+tibialis+raises+against+wall", "Back against wall, legs straight out. Lift toes upward toward shins as high as you can. Burn out for 15-20 reps.", "★ MICRO-LINK: Decelerates foot strike and balances calf volume."],
    ["Legs B", "Side Planks", "Quadratus Lumborum, Obliques, Outer Hip", "side-plank-leg-lift-hold-main.webp", "how+to+do+side+plank+proper+form", "Lie on side on forearm. Stack feet. Lift hips off ground until body forms a straight line from ears to heels. Squeeze glutes and brace obliques.", "Crucial lateral core endurance that shields lower back and pelvis."]
  ];

  var headers = [
    "Category",
    "Exercise Name",
    "Primary Muscle Group",
    "Visual Illustration",
    "Video Tutorial",
    "Form & Setup Cues (Plain English)",
    "Why It's in Your Program"
  ];

  var rows = [headers];

  for (var i = 0; i < data.length; i++) {
    var item = data[i];
    var category = item[0];
    var name = item[1];
    var muscles = item[2];
    var imgUrl = cdn + item[3];
    var videoUrl = "https://www.youtube.com/results?search_query=" + encodeURIComponent(item[4]);
    var cues = item[5];
    var why = item[6];

    var imgFormula = '=IMAGE("' + imgUrl + '", 1)';
    var videoFormula = '=HYPERLINK("' + videoUrl + '", "▶ Watch Video")';

    rows.push([
      category,
      name,
      muscles,
      imgFormula,
      videoFormula,
      cues,
      why
    ]);
  }

  sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);

  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 36);

  sheet.getRange(1, 1, 1, headers.length)
    .setBackground("#0f172a")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setFontSize(11)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  for (var r = 2; r <= rows.length; r++) {
    sheet.setRowHeight(r, 95);
    
    var isMicroLink = data[r - 2][2].indexOf("★") !== -1;
    var bg = isMicroLink ? "#f0fdf4" : ((r % 2 === 0) ? "#ffffff" : "#f8fafc");
    sheet.getRange(r, 1, 1, headers.length).setBackground(bg).setVerticalAlignment("middle");

    sheet.getRange(r, 1).setHorizontalAlignment("center").setFontWeight("bold");
    sheet.getRange(r, 2).setHorizontalAlignment("left").setFontWeight("bold").setFontSize(10);
    sheet.getRange(r, 3).setHorizontalAlignment("left").setFontSize(10);
    sheet.getRange(r, 4).setHorizontalAlignment("center");
    sheet.getRange(r, 5).setHorizontalAlignment("center").setFontColor("#2563eb").setFontWeight("bold");
    sheet.getRange(r, 6).setHorizontalAlignment("left").setWrap(true).setFontSize(9);
    sheet.getRange(r, 7).setHorizontalAlignment("left").setWrap(true).setFontColor("#475569").setFontSize(9);
  }

  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 240);
  sheet.setColumnWidth(3, 190);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 160);
  sheet.setColumnWidth(6, 420);
  sheet.setColumnWidth(7, 300);

  sheet.getRange(1, 1, rows.length, headers.length).setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);

  // Automatically update hyperlinks on the Workout Tracker tab
  try {
    addExerciseHyperlinks();
  } catch (err) {
    Logger.log("Note on hyperlinks: " + err.message);
  }
}

/**
 * Standalone function to add or refresh clickable hyperlinks in Column B
 * of the Workout Tracker. Each exercise will link directly to its row
 * in the Exercise Guide tab.
 */
function addExerciseHyperlinks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var guideSheet = ss.getSheetByName("Exercise Guide");
  
  if (!guideSheet) {
    SpreadsheetApp.getUi().alert("The 'Exercise Guide' tab was not found. Please run createExerciseGuideTab first.");
    return;
  }

  var guideGid = guideSheet.getSheetId();
  var guideData = guideSheet.getDataRange().getValues();
  
  // Build lookup map: category_name -> row, and name -> row
  var guideRowMap = {};
  for (var g = 1; g < guideData.length; g++) {
    var cat = guideData[g][0].toString().trim().toLowerCase();
    var name = guideData[g][1].toString().trim().toLowerCase();
    var rowNum = g + 1; // 1-based row in Google Sheets
    
    guideRowMap[name] = rowNum;
    guideRowMap[cat + "_" + name] = rowNum;
  }

  // Find Workout Tracker sheet (ID 717009429 or first main tab)
  var trackerSheet = null;
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === 717009429) {
      trackerSheet = sheets[i];
      break;
    }
  }
  if (!trackerSheet) {
    for (var j = 0; j < sheets.length; j++) {
      var sName = sheets[j].getName().toLowerCase();
      if (!sName.includes("copy") && !sName.includes("guide")) {
        trackerSheet = sheets[j];
        break;
      }
    }
  }
  if (!trackerSheet) trackerSheet = ss.getActiveSheet();

  var lastRow = trackerSheet.getLastRow();
  if (lastRow < 3) return;

  // Read Column A (Day) and Column B (Exercise) from row 3 downwards
  var valuesAB = trackerSheet.getRange(3, 1, lastRow - 2, 2).getValues();
  var rangeB = trackerSheet.getRange(3, 2, lastRow - 2, 1);
  var formulasB = rangeB.getFormulas();

  var linkedCount = 0;
  for (var r = 0; r < valuesAB.length; r++) {
    var day = valuesAB[r][0].toString().trim().toLowerCase();
    var rawName = valuesAB[r][1].toString().trim();

    // Skip blank cells or section dividers
    if (!rawName || rawName.indexOf("▶") === 0) continue;

    // Look up target row in Exercise Guide
    var cleanName = rawName.toLowerCase();
    var targetRow = guideRowMap[day + "_" + cleanName] || guideRowMap[cleanName];

    if (targetRow) {
      // Relative internal jump link to the exercise cell in Exercise Guide
      formulasB[r][0] = '=HYPERLINK("#gid=' + guideGid + '&range=B' + targetRow + '", "' + rawName + '")';
      linkedCount++;
    }
  }

  // Apply formulas back to Column B in a single batch
  rangeB.setFormulas(formulasB);

  // Set font line to underline and color links
  for (var r = 0; r < valuesAB.length; r++) {
    var rawName = valuesAB[r][1].toString().trim();
    if (rawName && rawName.indexOf("▶") !== 0) {
      var cell = trackerSheet.getRange(r + 3, 2);
      cell.setFontLine("underline");
      var fontColor = cell.getFontColor();
      if (fontColor !== "#15803d" && fontColor !== "#166534") {
        cell.setFontColor("#1d4ed8");
      }
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("Linked " + linkedCount + " exercises to the Exercise Guide!", "Links Updated", 5);
}
