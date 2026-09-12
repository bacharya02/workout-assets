/**
 * Comprehensive Mobility, Warmup & Stretching Guide Creator Script
 * 
 * Creates a dedicated "Mobility & Stretching" tab organized by:
 * 1. 🔵 PULL DAY (☀️ Pre-Workout Warmup + 🌙 Post-Workout Stretch)
 * 2. 🔴 PUSH DAY (☀️ Pre-Workout Warmup + 🌙 Post-Workout Stretch)
 * 3. 🟢 LEG DAY (☀️ Pre-Workout Warmup + 🌙 Post-Workout Stretch - Hip Protected)
 * 4. ⭐ REST DAY MASTER ROUTINE (Full-Body Hip Decompression & Mobility Flow)
 */

function createMobilityStretchingTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "Mobility & Stretching";
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

  // Array of routines with section divider flags
  // Format: [isDivider, routineDay, phase, name, target, timing, imgFile, videoQuery, cues, rationale]
  var rawData = [
    // =========================================================================
    // 🔵 PULL DAY ROUTINE
    // =========================================================================
    [true, "🔵 PULL DAY ROUTINE (Back, Lats, Biceps & Core)", "", "", "", "", "", "", "", ""],
    // Pull Pre-Workout
    [false, "Pull Day", "☀️ Pre-Workout Warmup", "Cat-Cow Spine Lubrication", "Lumbo-Pelvic Rhythm & Thoracic Spine", "10 slow breathing cycles", "cat-cow-main.webp", "how+to+cat+cow+proper+form", "Hands under shoulders, knees under hips. Inhale: gently drop belly, lift chest and tailbone (cow). Exhale: push floor away, dome upper back toward ceiling, tuck pelvis under (cat). Move vertebra by vertebra.", "Restores segmental spinal and pelvic mobility without any axial load or hip socket pinching."],
    [false, "Pull Day", "☀️ Pre-Workout Warmup", "Thoracic Spine Openers (Open Books)", "Mid-Back (T-Spine) Rotation & Ribcage Mobility", "8 reps/side (hold 2s)", "supine-spinal-twist-main.webp", "side+lying+open+book+stretch+thoracic", "Lie on side with knees bent 90° stacked together. Reach bottom arm forward, sweep top arm up and open across the ceiling toward the floor behind you, following hand with eyes. Keep knees glued together.", "If the mid-back is stiff, the hips and pelvis are forced to over-twist to compensate. Restoring T-spine rotation shields the hips."],
    [false, "Pull Day", "☀️ Pre-Workout Warmup", "Kneeling Wrist & Forearm Mobility Rocks", "Wrist Flexors, Extensors & Carpal Tunnel", "10 forward/backward pulses", "kneeling-wrist-stretch-main.webp", "wrist+mobility+routine+for+lifters", "On hands and knees, place palms flat with fingers pointing backward toward knees. Gently rock hips back to stretch forearms. Reverse with fingers pointing away and rock forward.", "Prepares the wrists and forearms for heavy pulling compounds with zero elbow or wrist strain."],
    [false, "Pull Day", "☀️ Pre-Workout Warmup", "Deadbugs (Deep Core Brace)", "Deep Core & Anterior Pelvic Stability", "8 reps/side (controlled)", "dead-bug-hold-main.webp", "how+to+do+deadbug+exercise", "Lie flat on back. Push lower back firmly into floor (no gap!). Arms straight up, knees at 90°. Slowly extend right arm overhead and left leg straight out. Return and alternate.", "Primes abdominal brace so heavy rows and pulldowns don't wrench the lumbar spine or pelvis."],
    // Pull Post-Workout
    [false, "Pull Day", "🌙 Post-Workout Stretch", "Banded / Pole Lat Stretch", "Lats (Latissimus Dorsi) & Outer Ribcage", "45 seconds/side", "banded-lat-stretch-main.webp", "banded+lat+stretch+proper+form", "Hold heavy resistance band (or upright gym rack) at head height with one hand. Step back, hinge at hips with soft knees, and let the band gently pull your arm and torso long. Breathe into the side body.", "The lat attaches directly to the thoracolumbar fascia on the pelvis. Tight lats pull the pelvis into anterior tilt; lengthening them relieves hip and lower back strain."],
    [false, "Pull Day", "🌙 Post-Workout Stretch", "Wide-Knee Bench Child’s Pose", "Latissimus Dorsi & Lumbar Decompression", "60 seconds (deep belly breaths)", "bench-childs-pose-main.webp", "bench+childs+pose+lat+stretch", "Kneel on floor with knees wide, placing hands/forearms on a bench in front. Sink hips back toward heels while dropping head and chest between arms. Inhale deeply expanding lower back ribs.", "Elevating hands on a bench maximizes lat decompression while wide knees allow the pelvis to drop without pinching the anterior hip capsule."],
    [false, "Pull Day", "🌙 Post-Workout Stretch", "Doorway Bicep & Chest Opener", "Biceps Brachii & Pectoralis Major", "45 seconds/side", "doorway-chest-stretch-main.webp", "doorway+bicep+and+chest+stretch", "Stand in doorway. Place palm and inside of forearm flat against doorframe at shoulder height. Step forward gently and turn torso away to stretch bicep and chest.", "Lengthens biceps and anterior shoulder after high-volume curls and rows."],
    [false, "Pull Day", "🌙 Post-Workout Stretch", "Neck & Upper Trapezius Release", "Upper Trapezius & Levator Scapulae", "30 seconds/side", "neck-side-stretch-main.webp", "upper+trap+and+neck+stretch+relief", "Sit upright. Gently drop right ear toward right shoulder. Reach left fingertips down toward floor (or sit on left palm) to anchor the shoulder down. Rest right hand on head for gentle weight only.", "De-escalates neck and upper shoulder tension accumulated from heavy rows, shrugs, and carries."],

    // =========================================================================
    // 🔴 PUSH DAY ROUTINE
    // =========================================================================
    [true, "🔴 PUSH DAY ROUTINE (Chest, Shoulders & Triceps)", "", "", "", "", "", "", "", ""],
    // Push Pre-Workout
    [false, "Push Day", "☀️ Pre-Workout Warmup", "Cat-Cow Spine Lubrication", "Thoracic Spine Extension & Scapular Motion", "10 slow cycles", "cat-cow-main.webp", "how+to+cat+cow+proper+form", "On hands and knees, smoothly alternate between arching mid-back (cow) and rounding upper back toward ceiling (cat). Focus on spreading shoulder blades wide at top.", "Frees up thoracic extension required for proper pressing angles without lower back arching."],
    [false, "Push Day", "☀️ Pre-Workout Warmup", "Shoulder External Rotations (Light Band)", "Rotator Cuff (Infraspinatus & Teres Minor)", "12-15 reps/side", "cable-external-rotation-peak.webp", "cable+shoulder+external+rotation+elbow+at+side", "Stand with elbow pinned tightly to side at 90°. Rotate forearm outward away from stomach against light band or cable. Smooth, controlled reps.", "Centrates the humeral head in the shoulder socket, preventing subacromial impingement during pressing."],
    [false, "Push Day", "☀️ Pre-Workout Warmup", "Push-Up Plus / Scapular Protraction", "Serratus Anterior & Scapular Stability", "12 reps (2s hold at top)", "archer-push-ups-peak.webp", "how+to+do+push+up+plus+for+serratus", "In a push-up or high plank position with arms locked straight. Push upper back toward ceiling by spreading shoulder blades wide apart. Hold 2s, sink back without bending elbows.", "Primes the serratus anterior to upwardly rotate the scapula during overhead and incline presses."],
    [false, "Push Day", "☀️ Pre-Workout Warmup", "Half-Kneeling Hip Flexor Rocks (Glute Tuck)", "Anterior Pelvic Stability & Hip Capsule", "8 rocks/side", "half-kneeling-hip-flexor-rock-peak.webp", "half+kneeling+hip+flexor+rock+glute+tuck", "Half-kneeling with back knee cushioned. Squeeze back glute hard and tuck tailbone under (posterior tilt). Rock forward just 2 inches. Feel stretch in front hip.", "Stabilizes the pelvis and lumbo-sacral junction so you don't hyperextend your injured hips on bench arch or overhead press."],
    // Push Post-Workout
    [false, "Push Day", "🌙 Post-Workout Stretch", "Doorway Chest & Anterior Delt Stretch", "Pectoralis Major & Anterior Deltoid", "45 seconds/side", "doorway-chest-stretch-main.webp", "doorway+pec+stretch+proper+form", "Stand in doorway. Forearm against frame at 90°, elbow at shoulder height. Step forward gently with one foot, rotate chest away. Clean chest stretch.", "Reverses forward rounded shoulder posture from benching; opens the anterior kinetic chain."],
    [false, "Push Day", "🌙 Post-Workout Stretch", "Overhead Triceps & Lat Stretch", "Triceps (Long Head) & Latissimus", "45 seconds/side", "banded-triceps-stretch-main.webp", "overhead+tricep+stretch+proper+form", "Reach one arm overhead, bend elbow to touch upper back. Use other hand to gently guide elbow backward. Keep ribs down and core braced.", "Restores resting length to the long head of the triceps after heavy extensions and dips."],
    [false, "Push Day", "🌙 Post-Workout Stretch", "Cross-Body Shoulder Stretch", "Posterior Deltoid & Posterior Capsule", "45 seconds/side", "cross-body-shoulder-stretch-main.webp", "cross+body+shoulder+stretch+proper+form", "Bring one arm straight across chest at shoulder height. Hook opposite forearm under elbow and gently pull arm closer to chest. Keep shoulder blade pulled down.", "Relieves posterior capsule tightness which can push the shoulder forward out of alignment."],
    [false, "Push Day", "🌙 Post-Workout Stretch", "Neck & Upper Trapezius Release", "Upper Trapezius & Levator Scapulae", "30 seconds/side", "neck-side-stretch-main.webp", "upper+trap+and+neck+stretch+relief", "Drop ear toward shoulder, reach opposite hand down to anchor shoulder blade. Gentle stretch only.", "Releases upper neck and shoulder tension accumulated from overhead pressing and lateral raises."],

    // =========================================================================
    // 🟢 LEG DAY ROUTINE (HIP-PROTECTED)
    // =========================================================================
    [true, "🟢 LEG DAY ROUTINE (Hip-Protected Quads, Glutes & Hamstrings)", "", "", "", "", "", "", "", ""],
    // Leg Pre-Workout
    [false, "Leg Day", "☀️ Pre-Workout Warmup", "Quadruped Wide-Knee Rockbacks", "Posterior Hip Capsule Decompression", "10 slow rocks", "childs-pose-main.webp", "quadruped+rockback+hip+mobility", "On hands and knees, take knees comfortably wide (outside hips) with big toes touching. Keep spine flat. Slowly glide hips back toward heels. Pause 2s, glide forward.", "CRITICAL FOR HIP: Wide knees clear the bony acetabular rim, allowing deep hip flexion without front groin pinching."],
    [false, "Leg Day", "☀️ Pre-Workout Warmup", "Glute Bridge Activation Pulses", "Gluteus Maximus & Hip Flexor Reciprocal Inhibition", "12-15 reps (2s peak squeeze)", "glute-bridge-peak.webp", "how+to+do+glute+bridges+proper+form", "Lie on back, knees bent, feet flat hip-width apart. Drive through heels to lift hips until thighs and torso form a straight line. Squeeze glutes aggressively at top. Pause 2s.", "Wakes up dormant glutes and creates 'reciprocal inhibition,' signaling chronically tight hip flexors to instantly relax."],
    [false, "Leg Day", "☀️ Pre-Workout Warmup", "Half-Kneeling Hip Flexor Rocks (Tucked Pelvis)", "Anterior Hip Capsule, Psoas & Rectus Femoris", "8-10 rocks/side", "half-kneeling-hip-flexor-rock-peak.webp", "half+kneeling+hip+flexor+rock+glute+tuck", "Half-kneeling with cushion under back knee. Squeeze kneeling glute hard and tuck tailbone under (posterior tilt). Gently rock forward 2-3 inches. Do NOT arch lower back.", "Tucking the pelvis stretches the psoas and anterior capsule safely without causing the femoral head to shear forward into the labrum."],
    [false, "Leg Day", "☀️ Pre-Workout Warmup", "Ankle Dorsiflexion Bench Rocks", "Talocrural Joint (Ankle Dorsiflexion) & Achilles", "10-12 rocks/side", "bench-ankle-stretch-main.webp", "ankle+dorsiflexion+mobility+rocks", "Place foot on bench or floor ~4 inches from wall. Keeping heel glued firmly down, drive knee forward over 2nd toe as far as possible without heel lifting. Hold 2s, return.", "Stiff ankles force compensatory hip internal rotation and knee valgus. Restoring ankle mobility protects the hip socket during gait."],
    [false, "Leg Day", "☀️ Pre-Workout Warmup", "Banded Clamshells (Rotator Cuff of Hip)", "Deep 6 Hip External Rotators & Glute Medius", "15 reps/side (2s hold)", "banded-clamshell-peak.webp", "how+to+do+clamshells+exercise+glutes", "Lie on side with loop band above knees. Knees bent 45°, feet glued together. Keeping feet touching, rotate top knee open toward ceiling without pelvis rolling back.", "Primes the deep external rotators (gemelli, obturators) to keep the ball centered in the socket before leg training."],
    // Leg Post-Workout
    [false, "Leg Day", "🌙 Post-Workout Stretch", "Figure-4 Glute Stretch (Supine on Back / Bench)", "Piriformis, Gluteus Medius & Deep Rotators", "45-60 seconds/side", "bench-figure-4-glute-stretch-main.webp", "figure+4+stretch+on+back+proper+form", "Lie on back (or sit on bench). Cross right ankle over left knee to make a '4'. Gently pull left thigh toward chest until stretch is felt in outer right glute. Neck relaxed.", "100x safer than Pigeon Pose! Pigeon crushes an injured labrum under full body weight. Figure-4 gives 100% control with zero groin impingement."],
    [false, "Leg Day", "🌙 Post-Workout Stretch", "Bench Couch Stretch (Glute Squeezed)", "Rectus Femoris, Psoas & Anterior Pelvic Tilt Relief", "45-60 seconds/side", "bench-couch-stretch-main.webp", "how+to+do+couch+stretch+without+arching+back", "Kneel on cushion with back foot resting up on bench/wall. Step other foot forward at 90°. Squeeze kneeling glute hard and pull belly button in to flatten lower back. Sit upright.", "Releases chronic anterior pelvic tilt and decompresses the hip joint by lengthening the two-joint rectus femoris muscle safely."],
    [false, "Leg Day", "🌙 Post-Workout Stretch", "Standing / Bench Calf & Achilles Stretch", "Gastrocnemius, Soleus & Plantar Fascia", "45 seconds/side", "standing-calf-stretch-main.webp", "proper+standing+calf+stretch+wall", "Face wall or bench. Step back leg straight back with heel pinned to floor and toes pointing straight ahead. Lean forward onto front knee while keeping back leg locked straight.", "Lengthens calves after high-volume walking and raises; restores ankle dorsiflexion which absorbs walking impact away from the hip."],
    [false, "Leg Day", "🌙 Post-Workout Stretch", "Gentle Seated Butterfly (Elevated Pelvis)", "Adductors (Inner Thigh / Groin) & Pelvic Floor", "60 seconds hold", "butterfly-stretch-main.webp", "seated+butterfly+stretch+proper+form", "Sit tall on yoga block or cushion. Soles of feet together, knees dropping outward. Hold ankles, keep chest tall. Allow gravity to open inner thighs naturally—never force down.", "Sitting on a cushion tilts the pelvis forward neutrally, allowing adductors to stretch without rounding the lower back or jamming hip joints."],

    // =========================================================================
    // ⭐ REST DAY FULL-BODY MASTER ROUTINE (15–20 Mins Hip Rehab & Decompression)
    // =========================================================================
    [true, "⭐ REST DAY FULL-BODY MASTER ROUTINE (15–20 Mins Hip Rehab & Decompression)", "", "", "", "", "", "", "", ""],
    [false, "Rest Day Master", "⭐ Step 1: Spinal Lubrication", "Cat-Cow Spine Lubrication", "Lumbo-Pelvic Rhythm & Full Spine", "10 slow breathing cycles", "cat-cow-main.webp", "how+to+cat+cow+proper+form", "Hands under shoulders, knees under hips. Smoothly flow between Cow (belly drops, chest lifts) and Cat (dome spine toward ceiling, tailbone tucks). Breathe deeply.", "Decompresses the entire spinal column and restores pelvic tilt control with zero joint stress."],
    [false, "Rest Day Master", "⭐ Step 2: Core Centration", "Bird-Dog Anti-Rotational Hold", "Contralateral Posterior Chain & Deep Core", "8 reps/side (3s hold at top)", "bird-dog-hold-main.webp", "how+to+bird+dog+exercise+proper+form", "Reach right arm forward and left heel straight back. Squeeze glute to lift leg. Keep hips and shoulders perfectly square to floor. Switch sides smoothly.", "Teaches the glute and opposite lat to stabilize the pelvis and spine without rotational torque."],
    [false, "Rest Day Master", "⭐ Step 3: Hip Socket Opener", "Quadruped Wide-Knee Hip Rockbacks", "Posterior Hip Capsule & Adductor Mobility", "12 smooth rocks (pause 2s at back)", "childs-pose-main.webp", "quadruped+rockback+hip+mobility", "Knees wide, big toes touching. Glide hips back toward heels. Feel space opening in the back of your hip joints. Stop before lower back rounds.", "Gently glides the femoral head posteriorly in the acetabular cup, creating room in the front joint crease."],
    [false, "Rest Day Master", "⭐ Step 4: Hip Flexor Reset", "Half-Kneeling Hip Flexor Rocks (Tucked)", "Anterior Capsule, Psoas & Hip Flexors", "10 rocks/side (glute squeezed)", "half-kneeling-hip-flexor-rock-peak.webp", "half+kneeling+hip+flexor+rock+glute+tuck", "Half-kneeling on cushion. Posteriorly tilt pelvis (tuck tailbone under like a scared dog). Gently glide forward 2 inches. Feel clean front hip stretch.", "Unlocks chronic hip flexor tightness from sitting without compressing the anterior labrum."],
    [false, "Rest Day Master", "⭐ Step 5: T-Spine Rotation", "Thoracic Spine Open Books (Side-Lying)", "Mid-Back (T-Spine) Rotation & Ribcage", "10 reps/side (slow)", "supine-spinal-twist-main.webp", "side+lying+open+book+stretch+thoracic", "Lie on side with knees bent 90° and stacked. Sweep top arm across ceiling to floor behind you. Follow hand with eyes. Keep knees glued together.", "Restores 3D thoracic rotation, preventing rotational compensation and shearing forces at the hip."],
    [false, "Rest Day Master", "⭐ Step 6: Glute & Piriformis", "Figure-4 Glute Stretch (Supine on Back)", "Piriformis, Gluteus Medius & Outer Hip", "60 seconds/side (gentle pull)", "bench-figure-4-glute-stretch-main.webp", "figure+4+stretch+on+back+proper+form", "Lie on back. Ankle crossed over opposite knee. Gently hug bottom thigh toward chest until clean outer glute stretch is achieved. Relax head and neck.", "Releases deep rotators that cause piriformis syndrome and sciatica with zero anterior impingement."],
    [false, "Rest Day Master", "⭐ Step 7: Quad & Psoas", "Bench / Wall Couch Stretch", "Rectus Femoris, Psoas & Quad Length", "60 seconds/side", "bench-couch-stretch-main.webp", "how+to+do+couch+stretch+without+arching+back", "Back foot propped up on bench/couch. Front foot forward. Squeeze rear glute and sit tall. Inhale into belly, exhale relaxing deeper into stretch.", "Restores true tissue length to the quads and hip flexors, neutralizing anterior pelvic tilt."],
    [false, "Rest Day Master", "⭐ Step 8: Lat & Spine Relief", "Wide-Knee Bench Child’s Pose", "Latissimus Dorsi & Lumbar Decompression", "90 seconds (deep belly breathing)", "bench-childs-pose-main.webp", "bench+childs+pose+lat+stretch", "Knees wide, arms on bench, sink hips to heels. Drop head between arms. Inhale deeply through nose expanding lower ribs, exhale letting spine melt down.", "Decompresses the entire posterior kinetic chain and down-regulates the central nervous system."],
    [false, "Rest Day Master", "⭐ Step 9: Chest & Posture", "Doorway Chest & Shoulder Opener", "Pectoralis Major & Anterior Deltoid", "45 seconds/side", "doorway-chest-stretch-main.webp", "doorway+pec+stretch+proper+form", "Forearm against doorframe at shoulder height. Step through and turn away. Breathe into chest.", "Opens tight chest tissues, reverses rounded shoulders, and improves diaphragm breathing."],
    [false, "Rest Day Master", "⭐ Step 10: Ankle & Achilles", "Standing / Bench Calf & Ankle Stretch", "Gastrocnemius, Soleus & Plantar Fascia", "45 seconds/side", "standing-calf-stretch-main.webp", "proper+standing+calf+stretch+wall", "Heel flat on floor, back leg straight. Lean forward into wall/bench until deep calf stretch is felt.", "Restores ankle dorsiflexion, ensuring your feet and ankles absorb ground impact instead of your hips."]
  ];

  var headers = [
    "Routine Day",
    "Phase / Purpose",
    "Movement Name",
    "Target Anatomy & Focus",
    "Prescription / Timing",
    "Visual Illustration",
    "Video Tutorial",
    "Form & Setup Cues (Plain English)",
    "Hip-Safe Biomechanical Rationale"
  ];

  var rows = [headers];
  var dividerRowIndices = [];

  for (var i = 0; i < rawData.length; i++) {
    var item = rawData[i];
    var isDivider = item[0];
    
    if (isDivider) {
      dividerRowIndices.push(rows.length + 1); // 1-based row
      var divRow = [item[1], "", "", "", "", "", "", "", ""];
      rows.push(divRow);
    } else {
      var routineDay = item[1];
      var phase = item[2];
      var name = item[3];
      var target = item[4];
      var timing = item[5];
      var imgUrl = cdn + item[6];
      var videoUrl = "https://www.youtube.com/results?search_query=" + encodeURIComponent(item[7]);
      var cues = item[8];
      var rationale = item[9];

      var imgFormula = '=IMAGE("' + imgUrl + '", 1)';
      var videoFormula = '=HYPERLINK("' + videoUrl + '", "▶ Watch Video")';

      rows.push([
        routineDay,
        phase,
        name,
        target,
        timing,
        imgFormula,
        videoFormula,
        cues,
        rationale
      ]);
    }
  }

  sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);

  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 36);

  // Header styling
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground("#0f172a")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setFontSize(11)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  // Style data rows & dividers
  for (var r = 2; r <= rows.length; r++) {
    if (dividerRowIndices.indexOf(r) !== -1) {
      // Divider row
      sheet.setRowHeight(r, 32);
      var divTitle = rows[r - 1][0];
      var divBg = "#0f172a";
      var divColor = "#38bdf8";
      
      if (divTitle.indexOf("PULL") !== -1) {
        divBg = "#1e3a8a"; divColor = "#93c5fd";
      } else if (divTitle.indexOf("PUSH") !== -1) {
        divBg = "#581c87"; divColor = "#d8b4fe";
      } else if (divTitle.indexOf("LEG") !== -1) {
        divBg = "#064e3b"; divColor = "#6ee7b7";
      } else if (divTitle.indexOf("REST") !== -1) {
        divBg = "#78350f"; divColor = "#fde68a";
      }

      sheet.getRange(r, 1, 1, headers.length).setBackground(divBg).setFontColor(divColor).setFontWeight("bold").setFontSize(11).setVerticalAlignment("middle");
      sheet.getRange(r, 1).setHorizontalAlignment("left");
    } else {
      // Movement row
      sheet.setRowHeight(r, 95);
      var routineDay = rows[r - 1][0];
      var phase = rows[r - 1][1];
      var isWarmup = phase.indexOf("Warmup") !== -1 || phase.indexOf("Step") !== -1;
      
      var bg = (r % 2 === 0) ? "#ffffff" : "#f8fafc";
      sheet.getRange(r, 1, 1, headers.length).setBackground(bg).setVerticalAlignment("middle");

      sheet.getRange(r, 1).setHorizontalAlignment("center").setFontWeight("bold").setFontSize(10);
      sheet.getRange(r, 2).setHorizontalAlignment("center").setFontWeight("bold").setFontSize(9).setFontColor(isWarmup ? "#15803d" : "#7c3aed");
      sheet.getRange(r, 3).setHorizontalAlignment("left").setFontWeight("bold").setFontSize(10);
      sheet.getRange(r, 4).setHorizontalAlignment("left").setFontSize(10);
      sheet.getRange(r, 5).setHorizontalAlignment("center").setFontWeight("bold").setFontSize(10);
      sheet.getRange(r, 6).setHorizontalAlignment("center");
      sheet.getRange(r, 7).setHorizontalAlignment("center").setFontColor("#2563eb").setFontWeight("bold");
      sheet.getRange(r, 8).setHorizontalAlignment("left").setWrap(true).setFontSize(9);
      sheet.getRange(r, 9).setHorizontalAlignment("left").setWrap(true).setFontColor("#475569").setFontSize(9);
    }
  }

  // Column Widths
  sheet.setColumnWidth(1, 130); // Routine Day
  sheet.setColumnWidth(2, 170); // Phase / Purpose
  sheet.setColumnWidth(3, 230); // Movement Name
  sheet.setColumnWidth(4, 230); // Target Anatomy & Focus
  sheet.setColumnWidth(5, 150); // Prescription / Timing
  sheet.setColumnWidth(6, 130); // Visual Illustration
  sheet.setColumnWidth(7, 150); // Video Tutorial
  sheet.setColumnWidth(8, 380); // Form & Setup Cues
  sheet.setColumnWidth(9, 340); // Hip-Safe Biomechanical Rationale

  sheet.getRange(1, 1, rows.length, headers.length).setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);

  SpreadsheetApp.getActiveSpreadsheet().toast("Mobility & Stretching routines created successfully!", "Complete", 5);
}
