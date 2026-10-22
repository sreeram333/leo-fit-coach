import os
import sys
import shutil

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from playwright.sync_api import sync_playwright

import base64

BASE_DIR = r"d:\Leo\coach"
BRAIN_DIR = r"C:\Users\SreeramThangavel\.gemini\antigravity\brain\616b8bb3-4536-4129-99df-84a074c80008"

def get_base64_img(rel_path):
    full_path = os.path.join(BASE_DIR, rel_path)
    if not os.path.exists(full_path):
        return rel_path
    ext = os.path.splitext(rel_path)[1].lower().replace('.', '')
    mime = "image/jpeg" if ext in ["jpg", "jpeg"] else "image/gif" if ext == "gif" else "image/png"
    try:
        with open(full_path, "rb") as f:
            encoded = base64.b64encode(f.read()).decode("utf-8")
        return f"data:{mime};base64,{encoded}"
    except Exception as e:
        print(f"Error base64 encoding {rel_path}: {e}")
        return rel_path

# 28 Finalized Exercises Data
DAYS_DATA = [
    {
        "id": "day1",
        "day_num": 1,
        "title": "Day 1 — Upper Push",
        "subtitle": "Upper Chest Shelf, 3D Shoulders & Horseshoe Triceps",
        "duration": "~65 mins",
        "focus": "Upper chest clavicular shelf, lateral deltoid capping, and heavy triceps lockout.",
        "color": "#3b82f6",
        "badge_bg": "rgba(59, 130, 246, 0.15)",
        "badge_border": "rgba(59, 130, 246, 0.4)",
        "exercises": [
            {
                "num": 1,
                "name": "Seated Incline Chest Machine",
                "muscle": "Upper Chest (Clavicular Head Shelf)",
                "sets_reps": "3 Sets × 8–10 Reps",
                "rest": "90s Rest",
                "gif": "d1_s1_opt_c.gif",
                "activation": "Loads the upper clavicular fibers of the pectoralis major directly beneath the collarbones to create a dense, elevated upper chest shelf.",
                "setup": "Adjust seat height so the machine handles align level with your upper chest / clavicle. Plant feet flat with heels driven into floor.",
                "form": "Pinch shoulder blades against back pad. Press diagonally upward and forward, driving biceps toward collarbones. Squeeze for 1s at top without rolling shoulders.",
                "mistakes": "Setting seat too low (turns into a shoulder press) or flaring elbows at 90° (impinges rotator cuffs). Keep elbows tucked at 45°–60°.",
                "progression": "When you hit 10, 10, 10 reps with zero shoulder compensation, increase machine weight by +2.5 kg next week.",
                "video_url": "https://www.youtube.com/watch?v=LiDArz1R2NU",
                "video_creator": "FIT.nl",
                "video_title": "Incline Machine Chest Press | Proper Form Tutorial for Upper Chest"
            },
            {
                "num": 2,
                "name": "Flat Barbell Bench Press",
                "muscle": "Mid & Lower Chest (Sternal Pec Mass)",
                "sets_reps": "3 Sets × 8–10 Reps",
                "rest": "2 mins Rest",
                "gif": "d1_s2_opt_a.gif",
                "activation": "Primary compound builder for thick sternal chest density and overall horizontal pushing power.",
                "setup": "Lie on bench with eyes directly under the bar. Retract shoulder blades, slight arch in lower back, feet planted firmly for leg drive. Grip slightly wider than shoulders.",
                "form": "Unrack bar and pull shoulders down into back pockets. Lower bar under control to lower sternum/nipple line with forearms vertical. Press up explosively in slight arc.",
                "mistakes": "Bouncing bar off chest, flaring elbows perpendicular to torso, or lifting glutes off bench. Keep heels grounded and descent controlled.",
                "progression": "Hit 10, 10, 10 reps -> add +1.25 kg to +2.5 kg next week. If lower reps (8, 7, 6), stay at this weight and aim for +1 rep.",
                "video_url": "https://www.youtube.com/watch?v=vcBig73ojpE",
                "video_creator": "Jeff Nippard",
                "video_title": "How To Get A Huge Bench Press with Perfect Technique"
            },
            {
                "num": 3,
                "name": "Seated Dumbbell Lateral Raise",
                "muscle": "Side Shoulders (Lateral Deltoid Width & V-Taper)",
                "sets_reps": "4 Sets × 12–15 Reps",
                "rest": "60s Rest",
                "gif": "d1_s3_opt_b.gif",
                "activation": "Isolates the lateral deltoid to widen shoulder span, instantly exaggerating the V-taper down to a narrow waist.",
                "setup": "Sit on edge of a flat bench with chest tilted forward ~15°. Seated position completely eliminates hip swinging and leg cheating.",
                "form": "Raise dumbbells outward in the scapular plane (30° forward from torso line). Lead with elbows and keep pinkies slightly elevated. 2s slow descent.",
                "mistakes": "Using heavy dumbbells and shrugging with upper traps, or raising arms straight out sideways/behind torso. Stay in the scapular plane!",
                "progression": "Stay at current weight until hitting 15, 15, 15, 15 clean reps, then increase dumbbell weight by +1 kg per hand.",
                "video_url": "https://www.youtube.com/watch?v=4hTUCDUQaNA",
                "video_creator": "Renaissance Periodization",
                "video_title": "The Only Dumbbell Lateral Raise Guide You Need (For Side Delt Growth)"
            },
            {
                "num": 4,
                "name": "Seated Barbell Overhead Press",
                "muscle": "Front Shoulders (Anterior Deltoids & Core)",
                "sets_reps": "3 Sets × 8–10 Reps",
                "rest": "90s Rest",
                "gif": "d1_s4_opt_a.gif",
                "activation": "Develops thick anterior deltoids and vertical pressing stability, carrying over to throwing velocity.",
                "setup": "Sit on upright bench (75°–85° back angle). Grip barbell just outside shoulder width with palms facing forward.",
                "form": "Inhale deeply into diaphragm to brace core. Lower bar with control to chin/collarbone level. Press bar vertically overhead, moving head through window at top lockout.",
                "mistakes": "Excessive lower back hyperextension (compensating for tight lats/shoulders) or short-changing depth by stopping above forehead level.",
                "progression": "Hit 10, 10, 10 reps -> increase bar weight by +1.25 kg to +2.5 kg.",
                "video_url": "https://www.youtube.com/watch?v=NAtsHyowOXg",
                "video_creator": "Colossus Fitness",
                "video_title": "How to PROPERLY Overhead Press | Proper OHP Form for Muscle Gain"
            },
            {
                "num": 5,
                "name": "Seated Two-Arm Overhead DB Extension",
                "muscle": "Triceps (Long Head - 60% of Arm Mass)",
                "sets_reps": "3 Sets × 10–12 Reps",
                "rest": "75s Rest",
                "gif": "d1_s5_opt_b.gif",
                "activation": "Overhead arm elevation places the triceps long head into maximum stretch under load—crucial for overall upper arm thickness.",
                "setup": "Sit upright with back supported. Hold a single dumbbell overhead with both hands in diamond grip under top plate.",
                "form": "Keep upper arms vertical and pinned near ears. Lower dumbbell behind neck until deep triceps stretch is achieved. Press back up to lockout.",
                "mistakes": "Flaring elbows wide sideways or stopping before deep stretch. Keep upper arms pointing straight up throughout.",
                "progression": "Hit 12, 12, 12 reps -> bump dumbbell weight up by +2 kg.",
                "video_url": "https://www.youtube.com/watch?v=fYqswDVbJDg",
                "video_creator": "Scott Herman Fitness",
                "video_title": "HOW TO: Overhead Triceps Extension (BEST EXERCISE FOR HUGE TRICEPS) || PERFECT FORM"
            },
            {
                "num": 6,
                "name": "Cable Straight-Bar or Rope Pushdown",
                "muscle": "Triceps (Lateral Horseshoe Cut)",
                "sets_reps": "3 Sets × 12–15 Reps",
                "rest": "60s Rest",
                "gif": "d1_s6_opt_a.gif",
                "activation": "Directly targets the lateral head of the triceps for the distinct carved 'horseshoe' profile on the outer arm.",
                "setup": "High cable pulley attachment. Slight forward lean at hips, feet shoulder-width, knees unlocked.",
                "form": "Pin elbows against ribcage like door hinges. Push bar/rope straight down until arms are fully extended. Squeeze triceps hard for 1 full second at bottom.",
                "mistakes": "Letting elbows drift forward on ascent (turns into shoulder movement) or rocking torso for momentum. Pin elbows tight!",
                "progression": "Hit 15, 15, 15 reps -> drop pin 1 plate (+2.5 kg).",
                "video_url": "https://www.youtube.com/watch?v=_w-HpW70nSQ",
                "video_creator": "Scott Herman Fitness",
                "video_title": "HOW TO: Cable Triceps Pushdown || 3 Golden Rules (FOR GROWTH)"
            }
        ]
    },
    {
        "id": "day2",
        "day_num": 2,
        "title": "Day 2 — Athletic Lower A",
        "subtitle": "Quad Sweep, Calves, Deep Core & APT Reset",
        "duration": "~65 mins",
        "focus": "Pure knee extension without lower back shear, and deep core pelvic reset.",
        "color": "#10b981",
        "badge_bg": "rgba(16, 185, 129, 0.15)",
        "badge_border": "rgba(16, 185, 129, 0.4)",
        "exercises": [
            {
                "num": 1,
                "name": "Smith Machine Squat (Feet 4\" Forward)",
                "muscle": "Quads (Pure Knee Extension - Zero Back Shear)",
                "sets_reps": "3 Sets × 6–8 Reps",
                "rest": "2 mins Rest",
                "gif": "d2_s1_opt_a.gif",
                "activation": "Forces complete quad isolation. Eliminates lumbar spinal shear and anterior pelvic tilt compression by leaning back into the guided bar path.",
                "setup": "Rest bar across upper traps. Place feet 4 inches FORWARD from bar line, shoulder-width apart, toes turned 15° out.",
                "form": "Keep torso completely vertical. Sit down by flexing knees until thighs hit parallel to floor. Drive through midfoot and heels to stand.",
                "mistakes": "Placing feet directly under the bar (forces hip hinge and lower back strain). Feet MUST be 4 inches forward for pure quad drive!",
                "progression": "Hit 8, 8, 8 reps -> increase barbell weight by +2.5 kg next week.",
                "video_url": "https://www.youtube.com/watch?v=AHnX-aimA4E",
                "video_creator": "Scott Herman Fitness",
                "video_title": "How To: Smith Machine- Squat (Feet Forward Quad Isolation)"
            },
            {
                "num": 2,
                "name": "Stationary Bulgarian Split Squat",
                "muscle": "Quads & Glutes (Unilateral Single-Leg Drive)",
                "sets_reps": "3 Sets × 8–10 Reps / leg",
                "rest": "90s Rest between legs",
                "gif": "d2_s2_opt_a.gif",
                "activation": "Builds single-leg sprint propulsion, corrects left-right muscular imbalances, and stabilizes the pelvic ring.",
                "setup": "Place top of rear foot on bench behind you. Take a comfortable stride forward with front working leg. Torso tilted forward ~20°.",
                "form": "Hold dumbbells at sides. Drop rear knee straight toward floor until front thigh reaches parallel. Drive through front heel to stand.",
                "mistakes": "Lifting front heel off floor, or hyper-arching lower back. Keep front foot planted flat and core braced.",
                "progression": "Hit 10 reps on both legs across all 3 sets -> increase dumbbells by +1 kg per hand.",
                "video_url": "https://www.youtube.com/watch?v=hPlKPjohFS0",
                "video_creator": "Squat University",
                "video_title": "The PERFECT Bulgarian Split Squat (Avoid These Errors!)"
            },
            {
                "num": 3,
                "name": "Leg Press (Narrow & Low Foot Placement)",
                "muscle": "Quads (Vastus Lateralis Outer Sweep & Teardrop)",
                "sets_reps": "3 Sets × 10–12 Reps",
                "rest": "90s Rest",
                "gif": "d2_s3_opt_a.gif",
                "activation": "Extreme knee flexion angle directs tension exclusively into the outer quad sweep and teardrop above the knee.",
                "setup": "Sit deep into carriage seat with hips pressed down. Place feet LOW on the plate platform and NARROW (approx. 6 inches apart).",
                "form": "Release catches. Lower sled smoothly until knees hit 90° flexion. Drive through balls and mid-foot back up, stopping just before knee lockout.",
                "mistakes": "Allowing hips/glutes to tuck and lift off the seat pad (causes lumbar rounding). Keep glutes pinned to pad throughout.",
                "progression": "Hit 12, 12, 12 reps -> add +5 kg to +10 kg to the sled.",
                "video_url": "https://www.youtube.com/watch?v=oujca3_Shgw",
                "video_creator": "Scott Herman Fitness",
                "video_title": "How To: Leg Press (Cybex - Low & Narrow Quad Sweep)"
            },
            {
                "num": 4,
                "name": "Standing Calf Raise on Block with DB",
                "muscle": "Calves (Gastrocnemius - Upper Diamond & Spring)",
                "sets_reps": "4 Sets × 12–15 Reps",
                "rest": "60s Rest",
                "gif": "d2_s4_opt_a.gif",
                "activation": "Straight-knee calf extension isolates the gastrocnemius diamond head, building reactive sprint elasticity for cricket fielding.",
                "setup": "Balls of feet on block or plate edge. Hold dumbbell in one hand, other hand lightly touching rack for stability.",
                "form": "Lower heels into a deep deficit stretch. HOLD FOR A DEAD 2-SECOND PAUSE at bottom. Explode up onto big toes and hold 1s at top.",
                "mistakes": "Bouncing quickly using Achilles tendon rebound. The 2-second dead stop at the bottom is mandatory to force muscle recruitment!",
                "progression": "Hit 15 reps on all 4 sets with dead pauses -> increase dumbbell by +2 kg.",
                "video_url": "https://www.youtube.com/watch?v=-qsRtp_PbVM",
                "video_creator": "Jeff Nippard",
                "video_title": "How To FORCE YOUR CALVES To Grow With Smarter Training Methods (Standing Straight-Knee Pause)"
            },
            {
                "num": 5,
                "name": "Captain's Chair Straight Leg Raise",
                "muscle": "Deep Core & Lower Abs (APT Fix & Transverse Abdominis)",
                "sets_reps": "3 Sets × 12–15 Reps",
                "rest": "60s Rest",
                "gif": "d2_s5_opt_b.gif",
                "activation": "Directly strengthens the lower abdominal wall and transverse abdominis to tuck the pelvis back into neutral alignment, flattening belly bulge.",
                "setup": "Support body on captain's chair apparatus with forearms on pads, back pressed firmly against backrest.",
                "form": "Keep legs extended (micro-bend if tight). Raise legs to parallel or above. CRUCIAL CUE: At top, curl your pelvis upward toward ribs. 3s slow descent.",
                "mistakes": "Swinging legs with momentum or only flexing hip flexors without posterior pelvic tuck. Curl the pelvis up!",
                "progression": "Hit 15 clean reps without swinging -> add 2-second isometric freeze at top of each rep.",
                "video_url": "https://www.youtube.com/watch?v=2TC-VJ40XwA",
                "video_creator": "Andrew Coates",
                "video_title": "Captain's Chair Knee / Leg Raises | Ab Exercise Technique (Pelvic Curl Focus)"
            }
        ]
    },
    {
        "id": "day3",
        "day_num": 3,
        "title": "Day 3 — Upper Pull + Chicken Day",
        "subtitle": "V-Taper Lats, Mid-Back Posture & Biceps Peak",
        "duration": "~65 mins",
        "focus": "Lat width for dramatic V-taper, rear delts and rhomboids to fix desk slump, plus biceps peak hypertrophy.",
        "color": "#8b5cf6",
        "badge_bg": "rgba(139, 92, 246, 0.15)",
        "badge_border": "rgba(139, 92, 246, 0.4)",
        "exercises": [
            {
                "num": 1,
                "name": "Assisted Pull-Up Machine",
                "muscle": "Lats (Latissimus Dorsi V-Taper Width)",
                "sets_reps": "3 Sets × 8–10 Reps",
                "rest": "90s Rest",
                "gif": "d3_s1_opt_b.gif",
                "activation": "The gold-standard vertical pull for flaring the lats outward, creating wide upper back proportions that make the waist look tiny.",
                "setup": "Adjust counterbalance pin. Kneel on pad. Take wide overhand grip on outer handles.",
                "form": "FRONT TO CHEST ONLY. Pull collarbones toward handles by driving elbows down into back pockets. Squeeze lats hard at top, lower to full arm hang.",
                "mistakes": "Pulling behind neck (damages cervical spine and shoulder joints) or cutting range of motion short. Full stretch at bottom!",
                "progression": "Hit 10, 10, 10 reps -> REDUCE the counterweight assistance by 1 pin notch.",
                "video_url": "https://www.youtube.com/watch?v=yltjwI98GmI",
                "video_creator": "DrillFit Training",
                "video_title": "How to Use The Assisted Pull Up Machine (Front Chest to Bar Form)"
            },
            {
                "num": 2,
                "name": "Seated Cable Row (Neutral V-Bar Grip)",
                "muscle": "Mid-Back (Rhomboids, Lower Traps & Mid-Lats)",
                "sets_reps": "3 Sets × 8–10 Reps",
                "rest": "90s Rest",
                "gif": "d3_s2_opt_b.gif",
                "activation": "Pulls scapulae together and strengthens mid-back thickness to permanently reverse rounded desk shoulders.",
                "setup": "Sit tall on cable row bench with feet on footrests, knees slightly bent. Attach close-grip V-bar handle.",
                "form": "Sit upright with proud chest. Pull V-bar into navel. Drive elbows back behind torso, pinch shoulder blades together for 1s pause.",
                "mistakes": "Swinging torso backward and using lower back momentum to heave the weight. Keep torso stationary at 90°.",
                "progression": "Hit 10, 10, 10 reps -> drop pin 1 plate down.",
                "video_url": "https://www.youtube.com/watch?v=7o2oolbmzeI",
                "video_creator": "Scott Herman Fitness",
                "video_title": "How To: Seated Cable Low-Row || PERFECT FORM (V-Bar Navel Pull)"
            },
            {
                "num": 3,
                "name": "Reverse Pec Deck Machine (Rear Fly)",
                "muscle": "Rear Shoulders (Posterior Delts & Slouch Fix)",
                "sets_reps": "4 Sets × 12–15 Reps",
                "rest": "60s Rest",
                "gif": "d3_s3_opt_b.gif",
                "activation": "Builds 3D rear deltoid caps, corrects IT desk slouch, and shields the rotator cuff during fast bowling.",
                "setup": "Sit facing machine pad. Adjust seat so handles are at eye level. Grip handles with neutral or overhand grip.",
                "form": "Slight soft bend in elbows. Sweep handles outward in a wide expanding semi-circle. Pull with posterior delts until handles align with shoulders.",
                "mistakes": "Bending arms into a row, or shrugging with upper traps. Lead with back of shoulders in an expansive arc.",
                "progression": "Hit 15, 15, 15, 15 reps -> increase resistance by +2.5 kg.",
                "video_url": "https://www.youtube.com/watch?v=6yMdhi2DVao",
                "video_creator": "Mind Pump TV",
                "video_title": "How To Properly Use The Rear Delt Fly Machine (+ Traps Protraction Cue)"
            },
            {
                "num": 4,
                "name": "Dumbbell Pullover across Flat Bench",
                "muscle": "Lats & Serratus (Lat Stretch & Ribcage Expansion)",
                "sets_reps": "3 Sets × 12–15 Reps",
                "rest": "75s Rest",
                "gif": "d3_s4_opt_c.gif",
                "activation": "Loaded lat stretch across horizontal bench combined with intense serratus anterior activation (the 'boxer's muscle' framing the ribs).",
                "setup": "Lie perpendicularly across a flat bench with only upper back/shoulders supported. Hips slightly lower than bench.",
                "form": "Hold single dumbbell with both hands cupping upper plate. With slight elbow bend, lower dumbbell in an arc over head until deep lat stretch is felt. Pull back over face to collarbones.",
                "mistakes": "Bending elbows excessively (turns into tricep extension) or thrusting hips up. Keep hips steady to maximize lat elongation.",
                "progression": "Hit 15, 15, 15 reps -> increase dumbbell weight by +2 kg.",
                "video_url": "https://www.youtube.com/watch?v=UOjSq3uidqY",
                "video_creator": "mountaindog1 (John Meadows)",
                "video_title": "Exercise Index: Dumbbell Lat Pullover | GET A WIDE BACK"
            },
            {
                "num": 5,
                "name": "Seated Dumbbell Hammer Curl",
                "muscle": "Arm Thickness (Brachialis & Outer Forearms)",
                "sets_reps": "3 Sets × 10–12 Reps",
                "rest": "60s Rest",
                "gif": "d3_s5_opt_b.gif",
                "activation": "Targets the brachialis muscle beneath the biceps, pushing the bicep peak upward and adding visible thickness to the upper arm.",
                "setup": "Sit tall on upright bench with dumbbells held in neutral grip (palms facing each other like hammers).",
                "form": "Pin elbows to sides. Curl dumbbells upward without twisting wrists. Hold peak squeeze for 1 second, 2s eccentric descent.",
                "mistakes": "Swinging shoulders forward to gain momentum, or separating elbows from ribcage.",
                "progression": "Hit 12, 12, 12 reps -> upgrade to next dumbbell weight (+1 kg/hand).",
                "video_url": "https://www.youtube.com/watch?v=8XLxfXROrTo",
                "video_creator": "Scott Herman Fitness",
                "video_title": "HOW TO: Dumbbell Hammer Curl (BICEPS PEAK BUILDER!) || PERFECT FORM"
            },
            {
                "num": 6,
                "name": "Preacher Curl (DB or Barbell)",
                "muscle": "Biceps Peak (Biceps Brachii Short & Long Head)",
                "sets_reps": "3 Sets × 10–12 Reps",
                "rest": "60s Rest",
                "gif": "d3_s6_opt_b.gif",
                "activation": "Eliminates all shoulder assistance and isolates the biceps brachii for maximum peak contraction and arm fullness.",
                "setup": "Sit at preacher bench with armpits snug against top pad. Triceps rested flat against angled cushion.",
                "form": "Grip EZ-bar or dumbbells underhand. Curl weight upward toward chin. Lower weight under control, STOPPING 1 INCH BEFORE FULL LOCKOUT to safeguard bicep tendon.",
                "mistakes": "Hyperextending and locking out elbows forcefully under heavy load (injury risk). Keep tension on muscle bellies!",
                "progression": "Hit 12, 12, 12 reps -> add +1.25 kg to +2.5 kg.",
                "video_url": "https://www.youtube.com/watch?v=nbcgEmZ0Be4",
                "video_creator": "Scott Herman Fitness",
                "video_title": "3 Reasons You're Wasting Your Time With Preacher Curls! (3 FIXES!)"
            }
        ]
    },
    {
        "id": "day4",
        "day_num": 4,
        "title": "Day 4 — Athletic Lower B",
        "subtitle": "Hamstrings, Glute Drive & Cricket Sprint Recovery",
        "duration": "~65 mins",
        "focus": "Posterior chain power, glute-ham tie-in, and Soleus calf endurance with 48h recovery buffer before Saturday cricket.",
        "color": "#f59e0b",
        "badge_bg": "rgba(245, 158, 11, 0.15)",
        "badge_border": "rgba(245, 158, 11, 0.4)",
        "exercises": [
            {
                "num": 1,
                "name": "Barbell Romanian Deadlift (RDL)",
                "muscle": "Hamstrings & Glute-Ham Tie-In (Hip Hinge)",
                "sets_reps": "3 Sets × 8–10 Reps",
                "rest": "2 mins Rest",
                "gif": "d4_s1_opt_b.gif",
                "activation": "Builds posterior chain hamstring elasticity and glute hip-extension power. The #1 exercise to prevent sprint tears between cricket wickets.",
                "setup": "Stand tall holding barbell with shoulder-width overhand grip. Feet hip-width apart, knees unlocked in soft micro-bend.",
                "form": "Initiate by pushing hips directly backward toward wall behind you. Shave bar down shins until mid-shin and hamstrings are at maximum stretch. Drive hips forward through heels.",
                "mistakes": "Rounding lower spine or turning into a knee-dominant squat. Spine must stay locked like a tabletop; movement comes entirely from hip hinge!",
                "progression": "Hit 10, 10, 10 reps -> add +2.5 kg to barbell.",
                "video_url": "https://www.youtube.com/watch?v=3VXmecChYYM",
                "video_creator": "Max Euceda",
                "video_title": "How to do the ROMANIAN DEADLIFT! | 2 Minute Form Tutorial"
            },
            {
                "num": 2,
                "name": "High & Wide Stance Leg Press",
                "muscle": "Gluteus Maximus (Glute Power & APT Reversal)",
                "sets_reps": "3 Sets × 10–12 Reps",
                "rest": "90s Rest",
                "gif": "d4_s2_leg_press_glute.gif",
                "activation": "High and wide stance shifts load from quads straight to the gluteus maximus, building explosive hip extension and reversing anterior pelvic tilt.",
                "setup": "Sit back firmly into sled. Place feet at the VERY TOP EDGE of the platform, WIDE stance (outside shoulder-width), toes flared 20°–30° outward.",
                "form": "Lower sled until knees reach 90°. Drive forcefully through HEELS to return sled up without locking knees. Zero lower back strain.",
                "mistakes": "Driving through toes or letting knees collapse inward. Keep knees tracking over pinky toes and drive hard through heels!",
                "progression": "Hit 12, 12, 12 reps -> add +5 kg to +10 kg to sled.",
                "video_url": "https://www.youtube.com/watch?v=M_xoJ0iRLFc",
                "video_creator": "Physique Development",
                "video_title": "How to LEG PRESS for Glutes | Improve Your Technique & Grow More Muscle"
            },
            {
                "num": 3,
                "name": "Lying Hamstring Leg Curl Machine",
                "muscle": "Hamstrings (Knee Flexion & Sprint Resilience)",
                "sets_reps": "3 Sets × 10–12 Reps",
                "rest": "75s Rest",
                "gif": "d4_s3_opt_b.gif",
                "activation": "Direct knee-flexion hamstring isolation protects the biceps femoris during explosive sprints and sudden stops on turf.",
                "setup": "Lie face down with roller pad against lower calves/Achilles tendon. Grasp machine handles.",
                "form": "GLUE HIPS FIRMLY INTO THE PAD. Curl heels directly toward glutes. Squeeze hamstrings hard for 1s at top, lower slowly over 2–3s.",
                "mistakes": "Allowing hips to rise up off bench to cheat the weight (strains lower back). Hips must stay glued flat!",
                "progression": "Hit 12, 12, 12 reps -> drop pin 1 plate (+2.5 kg).",
                "video_url": "https://www.youtube.com/watch?v=1Tq3QdYUuHs",
                "video_creator": "Scott Herman Fitness",
                "video_title": "How To: Prone Leg Curl (Cybex - Hips Glued to Pad)"
            },
            {
                "num": 4,
                "name": "Seated Calf Raise (Weights on Knees)",
                "muscle": "Calves (Soleus Muscle - Anti-Cramp Sprint Engine)",
                "sets_reps": "3 Sets × 12–15 Reps",
                "rest": "60s Rest",
                "gif": "d4_s4_seated_calf.gif",
                "activation": "With knees bent at 90°, the gastrocnemius is completely deactivated, directing 100% load onto the Soleus—the deep postural muscle that absorbs cricket sprint impact and eliminates cramps.",
                "setup": "Sit on bench, balls of feet elevated on a 3-inch block. Place heavy barbell or pair of heavy dumbbells across knees with padding.",
                "form": "Lower heels into full deficit stretch. Hold for 1 second. Drive up onto big toes and hold peak contraction for 1 second at top.",
                "mistakes": "Doing shallow bounces with light weight. The soleus thrives under heavy, deliberate, full-stretch repetitions.",
                "progression": "Hit 15, 15, 15 reps -> add +2.5 kg onto knees.",
                "video_url": "https://www.youtube.com/watch?v=fFWpWJy8ybU",
                "video_creator": "Live Lean TV",
                "video_title": "How To: Dumbbell Seated Calf Raise (90-Degree Knee Soleus Isolation)"
            },
            {
                "num": 5,
                "name": "Barbell Over-Bench Plate Pinch / Hold",
                "muscle": "Forearms & Grip (Wrist Flexors & Isometric Hold)",
                "sets_reps": "3 Sets × 45–60s Hold",
                "rest": "60s Rest",
                "gif": "d4_s5_opt_c.gif",
                "activation": "Develops iron hand/wrist grip endurance and forearm vascularity, giving direct control and power for cricket bowling, throwing, and heavy lifts.",
                "setup": "Stand tall holding smooth bumper plates pinched between fingers and thumb, or holding a heavy barbell with palms forward.",
                "form": "Stand tall, shoulders rolled back, core braced. Hold plates/barbell isometrically for 45–60 seconds without letting grip slip.",
                "mistakes": "Using lifting straps or letting posture slump forward. Maintain proud chest and iron pinch.",
                "progression": "When you comfortably reach 60 seconds on all 3 sets, increase plate weight.",
                "video_url": "https://www.youtube.com/watch?v=LARw21BBiDk",
                "video_creator": "FIT.nl",
                "video_title": "Plate Pinch | Proper Form Tutorial for Grip Strength"
            }
        ]
    },
    {
        "id": "day5",
        "day_num": 5,
        "title": "Day 5 — Upper Proportions",
        "subtitle": "Golden Ratio Aesthetics: Delts, Arms & Upper Chest Touch-up",
        "duration": "~65 mins",
        "focus": "Deltoid capping, arm circumference, and upper chest density heading into the weekend.",
        "color": "#ec4899",
        "badge_bg": "rgba(236, 72, 153, 0.15)",
        "badge_border": "rgba(236, 72, 153, 0.4)",
        "exercises": [
            {
                "num": 1,
                "name": "Incline Chest Press Machine",
                "muscle": "Upper Chest (Clavicular Pec Density Touch-Up)",
                "sets_reps": "3 Sets × 8–10 Reps",
                "rest": "90s Rest",
                "gif": "d5_s1_opt_b.gif",
                "activation": "Second weekly stimulus for upper chest accelerates clavicular thickness, building the dense muscular collarbone shelf.",
                "setup": "Seat positioned so handles align with upper chest collarbones. Plant feet flat.",
                "form": "Press diagonally upward and forward. Squeeze upper pectorals hard at peak extension, lower with a 2-second controlled tempo.",
                "mistakes": "Letting elbows flare excessively or lifting shoulders off pad. Keep upper back pinned flat.",
                "progression": "Hit 10, 10, 10 reps -> increase machine weight by +2.5 kg.",
                "video_url": "https://www.youtube.com/watch?v=LiDArz1R2NU",
                "video_creator": "FIT.nl",
                "video_title": "Incline Machine Chest Press | Proper Form Tutorial for Upper Chest"
            },
            {
                "num": 2,
                "name": "Seated Dumbbell Lateral Raise (Drop-Set)",
                "muscle": "Side Shoulders (Lateral Delts 3D Capping)",
                "sets_reps": "4 Sets × 12–15 Reps (Drop-Set on 4th)",
                "rest": "60s Rest",
                "gif": "d5_s2_opt_c.gif",
                "activation": "Maximal lateral delt pump. Sets 1-3 build hypertrophy; the 4th set drop-set forces extreme blood flow and rounded 3D cannonball delts.",
                "setup": "Sit on flat bench, slight 15° forward lean. Dumbbells at sides.",
                "form": "Raise in scapular plane (30° forward). On set 4: complete 12–15 reps, immediately drop dumbbells, grab 30% lighter pair with zero rest and pump out 8–10 more reps!",
                "mistakes": "Swinging torso during the drop-set. Stay strictly seated and let side deltoids do 100% of the work.",
                "progression": "Increase base dumbbell weight once the first 3 sets reach 15 clean reps.",
                "video_url": "https://www.youtube.com/watch?v=4hTUCDUQaNA",
                "video_creator": "Renaissance Periodization",
                "video_title": "The Only Dumbbell Lateral Raise Guide You Need (For Side Delt Growth)"
            },
            {
                "num": 3,
                "name": "Reverse Pec Deck Machine",
                "muscle": "Rear Shoulders (Posterior Delts & Rotator Cuff)",
                "sets_reps": "3 Sets × 12–15 Reps",
                "rest": "60s Rest",
                "gif": "d5_s3_opt_a.gif",
                "activation": "Locks in posterior delt roundness and reinforces upper back posture before weekend rest.",
                "setup": "Face chest pad, handles at eye level, soft elbow bend.",
                "form": "Sweep arms outward in wide expanding semi-circle. Squeeze back of shoulders for 1 second, control return.",
                "mistakes": "Yanking with back momentum or shrugging traps. Maintain slow, smooth tempo.",
                "progression": "Hit 15, 15, 15 reps -> increase resistance by +2.5 kg.",
                "video_url": "https://www.youtube.com/watch?v=6yMdhi2DVao",
                "video_creator": "Mind Pump TV",
                "video_title": "How To Properly Use The Rear Delt Fly Machine (+ BONUS TIP)"
            },
            {
                "num": 4,
                "name": "Barbell Close-Grip Bench Press",
                "muscle": "Triceps (Medial/Lateral Heads & Heavy Lockout)",
                "sets_reps": "3 Sets × 10–12 Reps",
                "rest": "90s Rest",
                "gif": "d5_s4_opt_c.gif",
                "activation": "Heavy compound triceps pressing movement that overloads medial and lateral heads with chest/anterior delt assistance.",
                "setup": "Lie on bench. Grip bar shoulder-width apart (hands ~12–14 inches apart). Do not grip too narrow!",
                "form": "Tuck elbows close to ribs (~30° angle). Lower bar to lower sternum under control. Press up explosively and lock out triceps hard.",
                "mistakes": "Placing hands touching each other (causes severe wrist and elbow strain). Hands must be shoulder-width apart!",
                "progression": "Hit 12, 12, 12 reps -> add +2.5 kg to barbell.",
                "video_url": "https://www.youtube.com/watch?v=UYJsFzqdgK4",
                "video_creator": "Scott Herman Fitness",
                "video_title": "HOW TO: Close-Grip Bench Press (TRICEPS BUILDER) || PERFECT FORM"
            },
            {
                "num": 5,
                "name": "Standing Cable Bicep Curl",
                "muscle": "Biceps (Biceps Brachii Peak & Constant Tension)",
                "sets_reps": "3 Sets × 10–12 Reps",
                "rest": "60s Rest",
                "gif": "d5_s5_opt_a.gif",
                "activation": "Constant continuous cable tension through the entire range of motion maximizes bicep peak contraction and fullness.",
                "setup": "Attach straight or EZ bar to low pulley. Stand tall 1 step back from cable tower.",
                "form": "Pin elbows to sides. Curl bar upward toward chin. Squeeze biceps hard at peak contraction for 1 second. Lower under 2s control.",
                "mistakes": "Rocking torso or letting elbows drift forward to cheat the curl. Pin elbows tight against ribs.",
                "progression": "Hit 12, 12, 12 reps -> drop pin 1 plate down.",
                "video_url": "https://www.youtube.com/watch?v=d7friQusjF8",
                "video_creator": "Team Evolve",
                "video_title": "Standing Straight Bar Cable Bicep Curls (Tutorial + Tips)"
            },
            {
                "num": 6,
                "name": "Seated Barbell Palms-Up Wrist Curl",
                "muscle": "Forearms (Wrist Flexor Mass & Vascularity)",
                "sets_reps": "3 Sets × 15–20 Reps",
                "rest": "45s Rest",
                "gif": "d5_s6_opt_a.gif",
                "activation": "Direct wrist-flexor hypertrophy creates dense, thick, vascular forearms visible below t-shirts and shirts.",
                "setup": "Kneel next to flat bench with forearms resting flat across pad, wrists hanging just over the edge. Palms facing up.",
                "form": "Allow barbell to roll gently into fingers for deep forearm stretch. Curl fingers and wrists up toward ceiling. Hold peak flex for 1s.",
                "mistakes": "Lifting forearms off bench during curl. Keep forearms pinned flat to isolate the wrist flexors.",
                "progression": "Hit 20, 20, 20 reps -> increase bar weight by +1.25 kg to +2.5 kg.",
                "video_url": "https://www.youtube.com/watch?v=qMtmHwaCmYI",
                "video_creator": "Scott Herman Fitness",
                "video_title": "How To: Seated Wrist Curl (Forearm Flexor Mass)"
            }
        ]
    }
]

def generate_html():
    days_html = ""
    for d in DAYS_DATA:
        ex_html = ""
        for ex in d["exercises"]:
            gif_data_uri = get_base64_img(f"gifs/{ex['gif']}")
            ex_html += f"""
            <div class="exercise-card" data-day="{d['id']}" data-search="{ex['name'].lower()} {ex['muscle'].lower()}">
                <div class="card-inner">
                    <div class="gif-container">
                        <img src="{gif_data_uri}" alt="{ex['name']}" loading="lazy" class="exercise-gif">
                        <div class="gif-badge">Move #{ex['num']} • Motion Loop</div>
                        <a href="{ex.get('video_url', '#')}" target="_blank" rel="noopener noreferrer" class="gif-video-link" title="Click to open verified YouTube form tutorial">
                            <span class="play-icon">▶</span> Watch Form Video ({ex.get('video_creator', 'YouTube')})
                        </a>
                    </div>
                    <div class="details-container">
                        <div class="header-row">
                            <div class="title-wrap">
                                <span class="badge-num">#{ex['num']}</span>
                                <h3 class="exercise-title">{ex['name']}</h3>
                            </div>
                            <span class="muscle-badge">{ex['muscle']}</span>
                        </div>
                        
                        <div class="metric-tags">
                            <span class="tag tag-sets">{ex['sets_reps']}</span>
                            <span class="tag tag-rest">{ex['rest']}</span>
                            <span class="tag tag-prog">Double Progression</span>
                            <a href="{ex.get('video_url', '#')}" target="_blank" rel="noopener noreferrer" class="tag tag-video" title="Watch full video tutorial">
                                ▶ Video Tutorial ({ex.get('video_creator', 'YouTube')}) ↗
                            </a>
                        </div>

                        <div class="info-blocks">
                            <div class="info-item">
                                <div class="info-label"><span class="icon">🎯</span> Target Muscle & Activation</div>
                                <div class="info-desc">{ex['activation']}</div>
                            </div>
                            
                            <div class="info-item">
                                <div class="info-label"><span class="icon">📐</span> Setup, Stance & Grip</div>
                                <div class="info-desc highlight-box">{ex['setup']}</div>
                            </div>

                            <div class="info-item">
                                <div class="info-label"><span class="icon">💡</span> Form Cues & Execution</div>
                                <div class="info-desc">{ex['form']}</div>
                            </div>

                            <div class="info-item video-item">
                                <div class="info-label text-video"><span class="icon">🎥</span> Verified Tutorial Video & Biomechanics</div>
                                <div class="info-desc">
                                    <a href="{ex.get('video_url', '#')}" target="_blank" rel="noopener noreferrer" class="video-anchor">
                                        <span class="video-play-badge">▶ PLAY TUTORIAL</span>
                                        <span class="video-anchor-title">{ex.get('video_title', ex['name'])}</span>
                                        <span class="video-channel-badge">by {ex.get('video_creator', 'YouTube')} ↗</span>
                                    </a>
                                </div>
                            </div>

                            <div class="info-item mistake-item">
                                <div class="info-label text-danger"><span class="icon">⚠️</span> Common Mistakes to Avoid</div>
                                <div class="info-desc">{ex['mistakes']}</div>
                            </div>

                            <div class="info-item prog-item">
                                <div class="info-label text-success"><span class="icon">📈</span> Next Session Progression Target</div>
                                <div class="info-desc">{ex['progression']}</div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            """
        
        days_html += f"""
        <section id="{d['id']}" class="day-section">
            <div class="day-header" style="border-left: 4px solid {d['color']};">
                <div class="day-header-top">
                    <span class="day-pill" style="background: {d['badge_bg']}; border-color: {d['badge_border']}; color: {d['color']};">
                        DAY {d['day_num']}
                    </span>
                    <span class="duration-badge">⏱️ {d['duration']}</span>
                </div>
                <h2 class="day-title">{d['title']}</h2>
                <p class="day-subtitle">{d['subtitle']}</p>
                <div class="day-focus"><strong>Focus:</strong> {d.get('focus', d['subtitle'])}</div>
            </div>
            <div class="exercise-grid">
                {ex_html}
            </div>
        </section>
        """

    full_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Leo's Master Fitness Transformation Guide (Muscular & Shredded)</title>
    <style>
        :root {{
            --bg-base: #090d16;
            --bg-card: #131b2e;
            --bg-card-hover: #18233c;
            --border-card: rgba(255, 255, 255, 0.08);
            --text-primary: #f8fafc;
            --text-secondary: #94a3b8;
            --text-muted: #64748b;
            --accent-blue: #3b82f6;
            --accent-emerald: #10b981;
            --accent-cyan: #06b6d4;
            --accent-amber: #f59e0b;
            --accent-purple: #8b5cf6;
            --accent-rose: #ef4444;
            --radius-lg: 16px;
            --radius-md: 10px;
            --radius-sm: 6px;
        }}
        
        * {{
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            -webkit-tap-highlight-color: transparent;
        }}

        body {{
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            background-color: var(--bg-base);
            color: var(--text-primary);
            line-height: 1.5;
            padding-bottom: 80px;
        }}

        /* Sticky Navigation Bar */
        .top-nav {{
            position: sticky;
            top: 0;
            z-index: 100;
            background: rgba(9, 13, 22, 0.92);
            backdrop-filter: blur(14px);
            -webkit-backdrop-filter: blur(14px);
            border-bottom: 1px solid var(--border-card);
            padding: 12px 20px;
        }}

        .nav-inner {{
            max-width: 1200px;
            margin: 0 auto;
            display: flex;
            flex-direction: column;
            gap: 12px;
        }}

        .nav-brand-row {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 10px;
        }}

        .brand-logo {{
            font-size: 1.15rem;
            font-weight: 800;
            letter-spacing: 0.5px;
            background: linear-gradient(135deg, #38bdf8 0%, #10b981 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            display: flex;
            align-items: center;
            gap: 8px;
        }}

        .athlete-chip {{
            font-size: 0.78rem;
            color: var(--text-secondary);
            background: rgba(255, 255, 255, 0.05);
            padding: 4px 10px;
            border-radius: 9999px;
            border: 1px solid var(--border-card);
        }}

        .nav-controls {{
            display: flex;
            gap: 10px;
            flex-wrap: wrap;
            align-items: center;
        }}

        .search-bar {{
            flex: 1;
            min-width: 240px;
            position: relative;
        }}

        .search-input {{
            width: 100%;
            background: rgba(255, 255, 255, 0.05);
            border: 1px solid var(--border-card);
            border-radius: 9999px;
            padding: 8px 16px 8px 36px;
            font-size: 0.85rem;
            color: var(--text-primary);
            outline: none;
            transition: all 0.2s;
        }}

        .search-input:focus {{
            border-color: var(--accent-cyan);
            background: rgba(255, 255, 255, 0.08);
            box-shadow: 0 0 0 3px rgba(6, 182, 212, 0.15);
        }}

        .search-icon {{
            position: absolute;
            left: 12px;
            top: 50%;
            transform: translateY(-50%);
            font-size: 0.85rem;
            color: var(--text-muted);
        }}

        .pill-tabs {{
            display: flex;
            gap: 6px;
            overflow-x: auto;
            padding-bottom: 2px;
            scrollbar-width: none;
        }}

        .pill-tabs::-webkit-scrollbar {{
            display: none;
        }}

        .tab-btn {{
            background: rgba(255, 255, 255, 0.05);
            border: 1px solid var(--border-card);
            color: var(--text-secondary);
            padding: 6px 14px;
            border-radius: 9999px;
            font-size: 0.8rem;
            font-weight: 600;
            white-space: nowrap;
            cursor: pointer;
            transition: all 0.2s;
        }}

        .tab-btn:hover, .tab-btn.active {{
            background: var(--accent-cyan);
            color: #041019;
            border-color: var(--accent-cyan);
            font-weight: 700;
        }}

        /* Container */
        .main-container {{
            max-width: 1200px;
            margin: 24px auto;
            padding: 0 16px;
            display: flex;
            flex-direction: column;
            gap: 36px;
        }}

        /* Hero / Target Card */
        .hero-card {{
            background: linear-gradient(135deg, rgba(19, 27, 46, 0.95) 0%, rgba(15, 23, 42, 0.85) 100%);
            border: 1px solid var(--border-card);
            border-radius: var(--radius-lg);
            padding: 26px 28px;
            box-shadow: 0 20px 40px -15px rgba(0, 0, 0, 0.5);
        }}

        @media (max-width: 768px) {{
            .hero-card {{
                padding: 18px 16px;
            }}
        }}

        .hero-content h1 {{
            font-size: 1.8rem;
            font-weight: 800;
            margin-bottom: 8px;
            line-height: 1.2;
        }}

        .target-sub {{
            font-size: 0.95rem;
            color: var(--accent-cyan);
            font-weight: 600;
            margin-bottom: 16px;
        }}

        .metrics-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(130px, 1fr));
            gap: 10px;
            margin-bottom: 16px;
        }}

        .metric-box {{
            background: rgba(255, 255, 255, 0.03);
            border: 1px solid var(--border-card);
            border-radius: var(--radius-sm);
            padding: 10px;
            text-align: center;
        }}

        .metric-value {{
            font-size: 1.1rem;
            font-weight: 800;
            color: #38bdf8;
        }}

        .metric-label {{
            font-size: 0.7rem;
            color: var(--text-muted);
            text-transform: uppercase;
            font-weight: 600;
            letter-spacing: 0.5px;
        }}

        .hero-bullets {{
            list-style: none;
            display: flex;
            flex-direction: column;
            gap: 6px;
            font-size: 0.84rem;
            color: var(--text-secondary);
        }}

        .hero-bullets li::before {{
            content: "✓ ";
            color: var(--accent-emerald);
            font-weight: 800;
        }}

        /* Day Headers */
        .day-section {{
            margin-bottom: 24px;
        }}

        .day-header {{
            background: var(--bg-card);
            border: 1px solid var(--border-card);
            border-radius: var(--radius-md);
            padding: 16px 20px;
            margin-bottom: 16px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.2);
        }}

        .day-header-top {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 6px;
        }}

        .day-pill {{
            padding: 3px 10px;
            border-radius: 9999px;
            font-size: 0.75rem;
            font-weight: 800;
            border: 1px solid;
            letter-spacing: 0.5px;
        }}

        .duration-badge {{
            font-size: 0.75rem;
            color: var(--text-muted);
            font-weight: 600;
        }}

        .day-title {{
            font-size: 1.35rem;
            font-weight: 800;
            color: var(--text-primary);
        }}

        .day-subtitle {{
            font-size: 0.85rem;
            color: var(--accent-cyan);
            margin-bottom: 6px;
            font-weight: 600;
        }}

        .day-focus {{
            font-size: 0.8rem;
            color: var(--text-secondary);
        }}

        /* Exercise Grid & Cards */
        .exercise-grid {{
            display: flex;
            flex-direction: column;
            gap: 16px;
        }}

        .exercise-card {{
            background: var(--bg-card);
            border: 1px solid var(--border-card);
            border-radius: var(--radius-lg);
            overflow: hidden;
            transition: transform 0.2s, border-color 0.2s, box-shadow 0.2s;
        }}

        .exercise-card:hover {{
            border-color: rgba(6, 182, 212, 0.35);
            box-shadow: 0 8px 24px rgba(0, 0, 0, 0.35);
        }}

        .card-inner {{
            display: grid;
            grid-template-columns: 340px 1fr;
            min-height: 280px;
        }}

        @media (max-width: 860px) {{
            .card-inner {{
                grid-template-columns: 1fr;
            }}
        }}

        .gif-container {{
            background: #080c14;
            display: flex;
            align-items: center;
            justify-content: center;
            position: relative;
            padding: 12px;
            border-right: 1px solid var(--border-card);
        }}

        @media (max-width: 860px) {{
            .gif-container {{
                border-right: none;
                border-bottom: 1px solid var(--border-card);
                max-height: 280px;
            }}
        }}

        .exercise-gif {{
            max-width: 100%;
            max-height: 260px;
            border-radius: var(--radius-md);
            object-fit: contain;
            box-shadow: 0 4px 16px rgba(0, 0, 0, 0.4);
        }}

        .gif-badge {{
            position: absolute;
            top: 12px;
            left: 12px;
            background: rgba(19, 27, 46, 0.85);
            backdrop-filter: blur(6px);
            padding: 4px 8px;
            border-radius: 9999px;
            font-size: 0.7rem;
            font-weight: 700;
            color: var(--text-secondary);
            border: 1px solid var(--border-card);
        }}

        .details-container {{
            padding: 20px;
            display: flex;
            flex-direction: column;
            gap: 12px;
            justify-content: center;
        }}

        .header-row {{
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            flex-wrap: wrap;
            gap: 8px;
        }}

        .title-wrap {{
            display: flex;
            align-items: center;
            gap: 8px;
        }}

        .badge-num {{
            background: rgba(255, 255, 255, 0.08);
            color: #38bdf8;
            font-size: 0.75rem;
            font-weight: 800;
            padding: 2px 7px;
            border-radius: var(--radius-sm);
        }}

        .exercise-title {{
            font-size: 1.15rem;
            font-weight: 800;
            color: var(--text-primary);
        }}

        .muscle-badge {{
            background: rgba(6, 182, 212, 0.12);
            color: var(--accent-cyan);
            border: 1px solid rgba(6, 182, 212, 0.3);
            font-size: 0.74rem;
            font-weight: 700;
            padding: 3px 10px;
            border-radius: 9999px;
            white-space: nowrap;
        }}

        .metric-tags {{
            display: flex;
            flex-wrap: wrap;
            gap: 6px;
        }}

        .tag {{
            font-size: 0.72rem;
            font-weight: 700;
            padding: 3px 8px;
            border-radius: var(--radius-sm);
        }}

        .tag-sets {{
            background: rgba(16, 185, 129, 0.15);
            color: var(--accent-emerald);
            border: 1px solid rgba(16, 185, 129, 0.3);
        }}

        .tag-rest {{
            background: rgba(245, 158, 11, 0.12);
            color: var(--accent-amber);
            border: 1px solid rgba(245, 158, 11, 0.3);
        }}

        .tag-prog {{
            background: rgba(59, 130, 246, 0.12);
            color: #60a5fa;
            border: 1px solid rgba(59, 130, 246, 0.3);
        }}

        .info-blocks {{
            display: flex;
            flex-direction: column;
            gap: 8px;
            margin-top: 4px;
        }}

        .info-item {{
            font-size: 0.8rem;
            line-height: 1.45;
        }}

        .info-label {{
            font-weight: 700;
            color: var(--text-secondary);
            font-size: 0.75rem;
            margin-bottom: 2px;
            display: flex;
            align-items: center;
            gap: 4px;
        }}

        .info-desc {{
            color: #cbd5e1;
        }}

        .highlight-box {{
            background: rgba(6, 182, 212, 0.06);
            border-left: 3px solid var(--accent-cyan);
            padding: 4px 8px;
            border-radius: 0 var(--radius-sm) var(--radius-sm) 0;
            font-weight: 500;
        }}

        .mistake-item {{
            background: rgba(239, 68, 68, 0.05);
            border: 1px solid rgba(239, 68, 68, 0.2);
            border-radius: var(--radius-sm);
            padding: 6px 10px;
        }}

        .text-danger {{
            color: #f87171 !important;
        }}

        .prog-item {{
            background: rgba(16, 185, 129, 0.05);
            border: 1px solid rgba(16, 185, 129, 0.2);
            border-radius: var(--radius-sm);
            padding: 6px 10px;
        }}

        .text-success {{
            color: #34d399 !important;
        }}

        /* Section Cards */
        .content-card {{
            background: var(--bg-card);
            border: 1px solid var(--border-card);
            border-radius: var(--radius-lg);
            padding: 24px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.2);
        }}

        .content-card h2 {{
            font-size: 1.3rem;
            font-weight: 800;
            margin-bottom: 14px;
            color: var(--text-primary);
            display: flex;
            align-items: center;
            gap: 8px;
        }}

        /* Tables */
        .table-responsive {{
            overflow-x: auto;
            border-radius: var(--radius-md);
            border: 1px solid var(--border-card);
        }}

        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 0.82rem;
            text-align: left;
        }}

        th {{
            background: rgba(255, 255, 255, 0.04);
            color: var(--text-secondary);
            font-weight: 700;
            padding: 10px 14px;
            border-bottom: 1px solid var(--border-card);
        }}

        td {{
            padding: 10px 14px;
            border-bottom: 1px solid rgba(255, 255, 255, 0.04);
            color: #cbd5e1;
        }}

        tr:last-child td {{
            border-bottom: none;
        }}

        tr:hover td {{
            background: rgba(255, 255, 255, 0.02);
        }}

        /* Staples Grid */
        .staples-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
            gap: 12px;
            margin-top: 12px;
        }}

        .staple-box {{
            background: rgba(255, 255, 255, 0.03);
            border: 1px solid var(--border-card);
            border-radius: var(--radius-sm);
            padding: 12px;
            display: flex;
            flex-direction: column;
            gap: 4px;
        }}

        .staple-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            font-weight: 700;
            font-size: 0.85rem;
            color: var(--text-primary);
        }}

        .staple-qty {{
            font-size: 0.72rem;
            background: rgba(245, 158, 11, 0.15);
            color: var(--accent-amber);
            padding: 2px 6px;
            border-radius: 9999px;
            font-weight: 700;
        }}

        .staple-desc {{
            font-size: 0.78rem;
            color: var(--text-muted);
        }}

        /* Floating Back to Top */
        .btt-btn {{
            position: fixed;
            bottom: 20px;
            right: 20px;
            background: var(--accent-cyan);
            color: #041019;
            width: 42px;
            height: 42px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: 800;
            font-size: 1.1rem;
            border: none;
            cursor: pointer;
            box-shadow: 0 4px 14px rgba(6, 182, 212, 0.4);
            z-index: 90;
            opacity: 0.85;
            transition: opacity 0.2s, transform 0.2s;
        }}

        .btt-btn:hover {{
            opacity: 1;
            transform: scale(1.08);
        }}

        /* Print Media Styles for PDF */
        @media print {{
            body {{
                background: #ffffff !important;
                color: #0f172a !important;
            }}
            .top-nav, .btt-btn, .search-bar, .pill-tabs {{
                display: none !important;
            }}
            .main-container {{
                max-width: 100% !important;
                margin: 0 !important;
                padding: 0 !important;
            }}
            .hero-card, .content-card, .day-header, .exercise-card {{
                background: #ffffff !important;
                border: 1px solid #e2e8f0 !important;
                box-shadow: none !important;
                color: #0f172a !important;
                page-break-inside: avoid;
                break-inside: avoid;
            }}
            .exercise-card {{
                margin-bottom: 20px !important;
            }}
            .card-inner {{
                display: flex !important;
                flex-direction: row !important;
            }}
            .gif-container {{
                background: #f8fafc !important;
                border-right: 1px solid #e2e8f0 !important;
                width: 260px !important;
            }}
            .exercise-gif {{
                max-height: 200px !important;
            }}
            .exercise-title, .day-title, .hero-content h1 {{
                color: #0f172a !important;
            }}
            .info-desc, .day-focus, .hero-bullets {{
                color: #334155 !important;
            }}
            .highlight-box {{
                background: #f1f5f9 !important;
                border-left-color: #0284c7 !important;
            }}
            .mistake-item {{
                background: #fff1f2 !important;
                border-color: #fecdd3 !important;
            }}
            .prog-item {{
                background: #f0fdf4 !important;
                border-color: #bbf7d0 !important;
            }}
            .text-danger {{
                color: #b91c1c !important;
            }}
            .text-success {{
                color: #15803d !important;
            }}
            .muscle-badge {{
                background: #e0f2fe !important;
                color: #0369a1 !important;
                border-color: #bae6fd !important;
            }}
        }}

        /* Video Tutorial Buttons & Links */
        .gif-video-link {{
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 6px;
            margin-top: 10px;
            width: 100%;
            padding: 7px 12px;
            background: linear-gradient(135deg, rgba(239, 68, 68, 0.2) 0%, rgba(220, 38, 38, 0.35) 100%);
            border: 1px solid rgba(239, 68, 68, 0.5);
            border-radius: var(--radius-sm);
            color: #fee2e2;
            text-decoration: none;
            font-size: 0.76rem;
            font-weight: 700;
            transition: all 0.2s;
        }}

        .gif-video-link:hover {{
            background: linear-gradient(135deg, rgba(239, 68, 68, 0.4) 0%, rgba(220, 38, 38, 0.6) 100%);
            border-color: #f87171;
            transform: translateY(-1px);
            box-shadow: 0 4px 12px rgba(239, 68, 68, 0.3);
            color: #fff;
        }}

        .tag-video {{
            background: rgba(239, 68, 68, 0.15) !important;
            color: #fca5a5 !important;
            border: 1px solid rgba(239, 68, 68, 0.4) !important;
            text-decoration: none;
            display: inline-flex;
            align-items: center;
            gap: 4px;
            transition: all 0.2s;
        }}

        .tag-video:hover {{
            background: rgba(239, 68, 68, 0.35) !important;
            color: #fff !important;
            border-color: #ef4444 !important;
            transform: translateY(-1px);
        }}

        .video-item {{
            background: rgba(239, 68, 68, 0.05);
            border: 1px solid rgba(239, 68, 68, 0.25);
            border-radius: var(--radius-sm);
            padding: 8px 10px;
        }}

        .text-video {{
            color: #f87171 !important;
        }}

        .video-anchor {{
            display: flex;
            align-items: center;
            gap: 8px;
            text-decoration: none;
            color: #f1f5f9;
            flex-wrap: wrap;
            font-size: 0.82rem;
            margin-top: 3px;
            transition: color 0.2s;
        }}

        .video-anchor:hover {{
            color: #38bdf8;
        }}

        .video-play-badge {{
            background: #dc2626;
            color: #fff;
            font-size: 0.66rem;
            font-weight: 800;
            padding: 2px 7px;
            border-radius: 4px;
            letter-spacing: 0.5px;
        }}

        .video-anchor-title {{
            font-weight: 600;
        }}

        .video-channel-badge {{
            background: rgba(255, 255, 255, 0.08);
            color: #94a3b8;
            font-size: 0.72rem;
            padding: 2px 7px;
            border-radius: 4px;
        }}

    </style>
</head>
<body>

    <!-- Sticky Navigation -->
    <header class="top-nav">
        <div class="nav-inner">
            <div class="nav-brand-row">
                <div class="brand-logo">
                    <span>⚡ LEO FIT COACH</span>
                    <span style="font-weight: 400; font-size: 0.85rem; color: var(--text-muted);">| MASTER BLUEPRINT</span>
                </div>
                <div class="athlete-chip">
                    27yo • 180cm • Target: 69–70kg @ 9.5–10% BF • 2,250–2,350 kcal
                </div>
            </div>

            <div class="nav-controls">
                <div class="search-bar">
                    <span class="search-icon">🔍</span>
                    <input type="text" id="searchInput" class="search-input" placeholder="Search exercise or muscle (e.g. Quads, Bench, Lats)..." oninput="filterCards()">
                </div>
                <div class="pill-tabs">
                    <button class="tab-btn active" onclick="filterDay('all', this)">All (28 Moves)</button>
                    <button class="tab-btn" onclick="filterDay('day1', this)">Day 1 Push</button>
                    <button class="tab-btn" onclick="filterDay('day2', this)">Day 2 Lower A</button>
                    <button class="tab-btn" onclick="filterDay('day3', this)">Day 3 Pull</button>
                    <button class="tab-btn" onclick="filterDay('day4', this)">Day 4 Lower B</button>
                    <button class="tab-btn" onclick="filterDay('day5', this)">Day 5 Proportions</button>
                    <button class="tab-btn" onclick="scrollToId('weekly')">Schedule</button>
                    <button class="tab-btn" onclick="scrollToId('nutrition')">Nutrition</button>
                    <button class="tab-btn" onclick="scrollToId('posture')">APT Fix</button>
                </div>
            </div>
        </div>
    </header>

    <main class="main-container">

        <!-- Target Overview Banner -->
        <section class="hero-card">
            <div class="hero-content">
                <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 8px; margin-bottom: 12px;">
                    <span class="day-pill" style="background: rgba(6, 182, 212, 0.15); border-color: rgba(6, 182, 212, 0.4); color: var(--accent-cyan);">
                        🎯 4–5 MONTH TRANSFORMATION ROADMAP
                    </span>
                    <span class="duration-badge">Phase: Muscular & Shredded (~9.5%–10% BF)</span>
                </div>
                <h1>Muscular & Shredded Transformation Blueprint</h1>
                <div class="target-sub">Target Physique: ~69–70 kg @ ~9.5%–10% Body Fat (Razor-Sharp Definition)</div>
                
                <div class="metrics-grid">
                    <div class="metric-box">
                        <div class="metric-value">69–70 kg</div>
                        <div class="metric-label">Target Weight</div>
                    </div>
                    <div class="metric-box">
                        <div class="metric-value">9.5%–10%</div>
                        <div class="metric-label">Body Fat</div>
                    </div>
                    <div class="metric-box">
                        <div class="metric-value">29.5"–30"</div>
                        <div class="metric-label">Tapered Waist</div>
                    </div>
                    <div class="metric-box">
                        <div class="metric-value">145–155g</div>
                        <div class="metric-label">Daily Protein</div>
                    </div>
                </div>

                <ul class="hero-bullets">
                    <li><strong>Aesthetic Golden Ratio</strong>: 3D rounded deltoid caps, wide lat V-taper, and thick upper chest shelf.</li>
                    <li><strong>Razor-Sharp Conditioning</strong>: Deeply etched six-pack, visible serratus anterior, and tight obliques.</li>
                    <li><strong>Postural Correction</strong>: Anterior Pelvic Tilt (APT) reversed through daily vacuums & high-stance glute drive.</li>
                    <li><strong>Athletic Stamina</strong>: Explosive sprint speed & zero cramping across 2 hours of Saturday turf cricket.</li>
                </ul>
            </div>
        </section>

        <!-- 5-Day Workout Splits (All 28 Exercises) -->
        <div id="workoutSections">
            {days_html}
        </div>

        <!-- Weekly Alignment & Cricket Protocol -->
        <section id="weekly" class="content-card">
            <h2>📅 Weekly Schedule & Cricket Match Alignment</h2>
            <p style="font-size: 0.85rem; color: var(--text-secondary); margin-bottom: 14px;">
                Designed specifically to give Leo a full <strong>48-hour sprint recovery window</strong> between Thursday leg work and Saturday turf cricket.
            </p>
            <div class="table-responsive">
                <table>
                    <thead>
                        <tr>
                            <th>Day</th>
                            <th>Split</th>
                            <th>Focus Muscles</th>
                            <th>Recovery & Nutrition Status</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td><strong>Mon (Day 1)</strong></td>
                            <td>Upper Push</td>
                            <td>Upper Chest, 3D Lateral Delts, Triceps Lockout</td>
                            <td>Whey + 3.1g Creatine Post-Gym | 2,250–2,350 kcal</td>
                        </tr>
                        <tr>
                            <td><strong>Tue (Day 2)</strong></td>
                            <td>Athletic Lower A</td>
                            <td>Quads (Feet 4" Forward), Calves, Deep Core & APT Fix</td>
                            <td>High quad fatigue, zero spinal shear</td>
                        </tr>
                        <tr>
                            <td><strong>Wed (Day 3)</strong></td>
                            <td>Upper Pull</td>
                            <td>V-Taper Lats, Mid-Back Posture, Biceps Peak</td>
                            <td>🍗 <strong>Chicken Day</strong>: Home chicken gravy (No whey needed)</td>
                        </tr>
                        <tr>
                            <td><strong>Thu (Day 4)</strong></td>
                            <td>Athletic Lower B</td>
                            <td>Hamstrings, Glute Drive, Soleus Anti-Cramp Engine</td>
                            <td><strong>48h Buffer</strong> begins before Saturday cricket!</td>
                        </tr>
                        <tr>
                            <td><strong>Fri (Day 5)</strong></td>
                            <td>Upper Proportions</td>
                            <td>Lateral Delt Drop-Set, Arms, Upper Chest Touch-Up</td>
                            <td>Upper body pump; legs fully rested for match day</td>
                        </tr>
                        <tr>
                            <td><strong>Sat (Day 6)</strong></td>
                            <td>🏏 Turf Cricket</td>
                            <td>2-Hour High Intensity Sprinting & Fielding</td>
                            <td><strong>+50g Lunch Carbs</strong> + 1.5L pre-match hydration</td>
                        </tr>
                        <tr>
                            <td><strong>Sun (Day 7)</strong></td>
                            <td>Full Rest</td>
                            <td>Mobility, Psoas Stretch, 8,000 Step Stroll</td>
                            <td>🍗 <strong>Chicken Day</strong>: Full glycogen replenishment</td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </section>

        <!-- Nutrition & 8 Staples -->
        <section id="nutrition" class="content-card">
            <h2>🥗 Daily Nutrition Architecture & The 8 Non-Negotiable Staples</h2>
            <div class="metrics-grid" style="margin-bottom: 16px;">
                <div class="metric-box">
                    <div class="metric-value" style="color: #10b981;">2,250–2,350</div>
                    <div class="metric-label">Daily Calories (kcal)</div>
                </div>
                <div class="metric-box">
                    <div class="metric-value" style="color: #38bdf8;">145–155g</div>
                    <div class="metric-label">Protein Floor</div>
                </div>
                <div class="metric-box">
                    <div class="metric-value" style="color: #f59e0b;">65–75g</div>
                    <div class="metric-label">Healthy Fats</div>
                </div>
                <div class="metric-box">
                    <div class="metric-value" style="color: #8b5cf6;">250–275g</div>
                    <div class="metric-label">Complex Carbs</div>
                </div>
                <div class="metric-box">
                    <div class="metric-value" style="color: #06b6d4;">3.5–4.0L</div>
                    <div class="metric-label">Daily Hydration</div>
                </div>
                <div class="metric-box">
                    <div class="metric-value" style="color: #ec4899;">3.1g</div>
                    <div class="metric-label">Creatine Daily</div>
                </div>
            </div>

            <h3 style="font-size: 1rem; font-weight: 700; margin-bottom: 10px; color: var(--accent-cyan);">The 8 Non-Negotiable Daily Staples</h3>
            <div class="staples-grid">
                <div class="staple-box">
                    <div class="staple-header">
                        <span>1. Whole Eggs Only</span>
                        <span class="staple-qty">Up to 5 / day</span>
                    </div>
                    <div class="staple-desc">Never separate egg whites. Provides bioavailable protein, choline, lutein, and healthy cholesterol for testosterone.</div>
                </div>
                <div class="staple-box">
                    <div class="staple-header">
                        <span>2. Toned Milk</span>
                        <span class="staple-qty">Up to 500ml / day</span>
                    </div>
                    <div class="staple-desc">Split between morning oats and evening tea/coffee. Rich in casein and calcium for bone density.</div>
                </div>
                <div class="staple-box">
                    <div class="staple-header">
                        <span>3. Pottukadalai (Roasted Chana)</span>
                        <span class="staple-qty">50g / day</span>
                    </div>
                    <div class="staple-desc">Afternoon 4:30 PM crunch providing ~11g plant protein and 5g gut fiber without cooking effort.</div>
                </div>
                <div class="staple-box">
                    <div class="staple-header">
                        <span>4. Peanuts & Cashews</span>
                        <span class="staple-qty">15g + 15g daily</span>
                    </div>
                    <div class="staple-desc">15g roasted peanuts + 15g raw cashews. Delivers zinc, magnesium, and essential monounsaturated fats.</div>
                </div>
                <div class="staple-box">
                    <div class="staple-header">
                        <span>5. Pure Honey</span>
                        <span class="staple-qty">1–2 tsp (15g)</span>
                    </div>
                    <div class="staple-desc">Natural morning sweetener for oats or pre-gym fast glucose to top off liver glycogen.</div>
                </div>
                <div class="staple-box">
                    <div class="staple-header">
                        <span>6. Rolled / Instant Oats</span>
                        <span class="staple-qty">50g daily</span>
                    </div>
                    <div class="staple-desc">Low-glycemic complex carbohydrates and beta-glucan fiber to sustain morning mental and physical energy.</div>
                </div>
                <div class="staple-box">
                    <div class="staple-header">
                        <span>7. Soya Chunks Cap</span>
                        <span class="staple-qty">45g dry-weighed</span>
                    </div>
                    <div class="staple-desc">Delivers ~23.4g protein on non-chicken days. Must be weighed raw on a digital scale before boiling.</div>
                </div>
                <div class="staple-box">
                    <div class="staple-header">
                        <span>8. Chicken Gravy (Wed & Sun)</span>
                        <span class="staple-qty">Home-Cooked</span>
                    </div>
                    <div class="staple-desc">Fresh morning batch spread across breakfast, lunch, and dinner. No whey isolate powder needed on these days!</div>
                </div>
            </div>

            <h3 style="font-size: 1rem; font-weight: 700; margin: 20px 0 10px; color: var(--accent-amber);">Smart 1-to-1 Substitution Matrix (Missing Items)</h3>
            <div class="table-responsive">
                <table>
                    <thead>
                        <tr>
                            <th>Missing Staple</th>
                            <th>Smart 1-to-1 Replacement</th>
                            <th>Macros Matched</th>
                            <th>Action</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td><strong>Soya Chunks (45g)</strong></td>
                            <td><strong>100g pan-seared Tofu + 1 Boiled Egg</strong></td>
                            <td>~21.5g Protein</td>
                            <td>Sauté tofu cubes in 1/2 tsp oil with turmeric & salt.</td>
                        </tr>
                        <tr>
                            <td><strong>Both Soya & Tofu Out</strong></td>
                            <td><strong>150g Boiled Sundal (Black Chana) + 1 Egg</strong></td>
                            <td>~20.0g Protein</td>
                            <td>Boil black chana, temper with mustard seeds & curry leaves.</td>
                        </tr>
                        <tr>
                            <td><strong>Toned Milk (250ml)</strong></td>
                            <td><strong>200g Fresh Plain Curd (Dahi)</strong></td>
                            <td>~8.0g Protein</td>
                            <td>Enjoy with lunch rice or whisk into refreshing buttermilk.</td>
                        </tr>
                        <tr>
                            <td><strong>Pottukadalai (50g)</strong></td>
                            <td><strong>30g Boiled Peanuts or 2 slices bread + 1 egg</strong></td>
                            <td>~10.0g Protein</td>
                            <td>Quick toasted egg sandwich or salted boiled peanuts.</td>
                        </tr>
                        <tr>
                            <td><strong>Cashews / Peanuts</strong></td>
                            <td><strong>10–12 Raw Almonds or 15g Walnuts</strong></td>
                            <td>~3g P, 8g Fats</td>
                            <td>Soaked or raw with evening green tea/coffee.</td>
                        </tr>
                    </tbody>
                </table>
            </div>

            <h3 style="font-size: 1rem; font-weight: 700; margin: 20px 0 10px; color: var(--accent-emerald);">Dynamic Nutrition Rebalancing Protocols</h3>
            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 12px;">
                <div style="background: rgba(255, 255, 255, 0.03); border: 1px solid var(--border-card); border-radius: var(--radius-sm); padding: 12px;">
                    <strong style="color: var(--accent-cyan);">Scenario 1: Missed / Skipped Meal</strong>
                    <p style="font-size: 0.8rem; color: var(--text-secondary); margin-top: 4px;">
                        Do not skip protein. Roll the missed protein into evening or dinner: add 2 whole boiled eggs, extra 150ml milk, or 100g curd. Always protect the 145g protein floor!
                    </p>
                </div>
                <div style="background: rgba(255, 255, 255, 0.03); border: 1px solid var(--border-card); border-radius: var(--radius-sm); padding: 12px;">
                    <strong style="color: var(--accent-amber);">Scenario 2: Heavy / Unplanned Meal</strong>
                    <p style="font-size: 0.8rem; color: var(--text-secondary); margin-top: 4px;">
                        Do not starve. For remaining meals: drop 1 chapati, skip peanuts, remove cooking oil, and take pure protein (boiled eggs or whey in water). Add a 20-min evening stroll (+2,500 steps).
                    </p>
                </div>
            </div>
        </section>

        <!-- Morning 5-Minute APT & Posture Reset -->
        <section id="posture" class="content-card">
            <h2>🧘 Daily 5-Minute Morning Posture & APT Reset Protocol</h2>
            <p style="font-size: 0.85rem; color: var(--text-secondary); margin-bottom: 14px;">
                Perform every morning immediately after using the washroom and drinking 500ml room-temperature water.
            </p>
            <div style="display: flex; flex-direction: column; gap: 10px;">
                <div style="background: rgba(6, 182, 212, 0.06); border-left: 3px solid var(--accent-cyan); padding: 12px; border-radius: 0 var(--radius-sm) var(--radius-sm) 0;">
                    <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 6px;">
                        <strong style="color: var(--accent-cyan);">1. Stomach Vacuums (Transverse Abdominis Tucking)</strong>
                        <a href="https://www.youtube.com/watch?v=b5GsE6ZHn_E" target="_blank" rel="noopener noreferrer" class="tag tag-video">▶ Mind Pump TV Tutorial ↗</a>
                    </div>
                    <div style="font-size: 0.8rem; color: #cbd5e1; margin-top: 4px;">
                        <strong>4 sets × 20-second isometric holds</strong>. Exhale all air completely, pull belly button back toward spine as hard as possible, and hold breath/vacuum. Trains the inner anatomical corset to suck in lower abdominal protrusion.
                    </div>
                </div>
                <div style="background: rgba(16, 185, 129, 0.06); border-left: 3px solid var(--accent-emerald); padding: 12px; border-radius: 0 var(--radius-sm) var(--radius-sm) 0;">
                    <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 6px;">
                        <strong style="color: var(--accent-emerald);">2. Kneeling Hip Flexor Stretch with Posterior Pelvic Tilt</strong>
                        <a href="https://www.youtube.com/watch?v=gqoPYLUgP48" target="_blank" rel="noopener noreferrer" class="tag tag-video">▶ The Prehab Guys Tutorial ↗</a>
                    </div>
                    <div style="font-size: 0.8rem; color: #cbd5e1; margin-top: 4px;">
                        <strong>2 sets × 45–60s per leg</strong>. Kneel on one knee, squeeze the rear glute hard to tuck pelvis backward, and shift forward slightly until front hip flexor (psoas) stretches. Reverses hours of IT desk chair sitting.
                    </div>
                </div>
                <div style="background: rgba(245, 158, 11, 0.06); border-left: 3px solid var(--accent-amber); padding: 12px; border-radius: 0 var(--radius-sm) var(--radius-sm) 0;">
                    <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 6px;">
                        <strong style="color: var(--accent-amber);">3. Glute Squeezes & Pelvic Tucks</strong>
                        <a href="https://www.youtube.com/watch?v=mkbNJKiWNlI" target="_blank" rel="noopener noreferrer" class="tag tag-video">▶ Athlean-X APT Fix Guide ↗</a>
                    </div>
                    <div style="font-size: 0.8rem; color: #cbd5e1; margin-top: 4px;">
                        <strong>2 sets × 10 reps (3-second squeeze at top)</strong>. Lie on back or stand tall. Squeeze gluteus maximus to tuck tailbone under, waking up glute motor units before sitting at desk.
                    </div>
                </div>
            </div>
        </section>

    </main>

    <!-- Floating Back to Top Button -->
    <button class="btt-btn" onclick="window.scrollTo({{ top: 0, behavior: 'smooth' }})">↑</button>

    <script>
        function filterDay(dayId, btn) {{
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            if (btn) btn.classList.add('active');

            const sections = document.querySelectorAll('.day-section');
            sections.forEach(sec => {{
                if (dayId === 'all') {{
                    sec.style.display = 'block';
                }} else {{
                    if (sec.id === dayId) {{
                        sec.style.display = 'block';
                        sec.scrollIntoView({{ behavior: 'smooth', block: 'start' }});
                    }} else {{
                        sec.style.display = 'none';
                    }}
                }}
            }});
        }}

        function filterCards() {{
            const query = document.getElementById('searchInput').value.toLowerCase().trim();
            const cards = document.querySelectorAll('.exercise-card');
            cards.forEach(card => {{
                const searchData = card.getAttribute('data-search') || '';
                if (!query || searchData.includes(query)) {{
                    card.style.display = 'block';
                }} else {{
                    card.style.display = 'none';
                }}
            }});
        }}

        function scrollToId(id) {{
            const el = document.getElementById(id);
            if (el) {{
                el.scrollIntoView({{ behavior: 'smooth', block: 'start' }});
            }}
        }}
    </script>
</body>
</html>
"""
    return full_html

def build_pdf(html_path, pdf_path):
    print(f"Generating PDF via Playwright: {pdf_path}")
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()
        page.goto(f"file:///{os.path.abspath(html_path).replace(os.sep, '/')}")
        page.wait_for_timeout(1000)
        page.pdf(
            path=pdf_path,
            format="A4",
            print_background=True,
            margin={"top": "12mm", "bottom": "12mm", "left": "10mm", "right": "10mm"}
        )
        browser.close()
    print("PDF generation complete!")

def build_docx(docx_path):
    print(f"Generating DOCX via python-docx: {docx_path}")
    doc = Document()
    
    # Page setup
    for section in doc.sections:
        section.top_margin = Inches(0.7)
        section.bottom_margin = Inches(0.7)
        section.left_margin = Inches(0.7)
        section.right_margin = Inches(0.7)
        
    # Styles helper
    def set_cell_bg(cell, fill_hex):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
        tcPr.append(shd)

    # Document Header
    p_title = doc.add_paragraph()
    run_title = p_title.add_run("LEO'S MASTER FITNESS TRANSFORMATION BLUEPRINT")
    run_title.font.name = "Calibri"
    run_title.font.size = Pt(22)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(14, 116, 144) # Dark cyan
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p_sub = doc.add_paragraph()
    run_sub = p_sub.add_run("Target Outcome: Muscular & Shredded (~9.5%–10% Body Fat, ~69–70 kg, Etched 6-Pack, Cricket Sprint Power)")
    run_sub.font.name = "Calibri"
    run_sub.font.size = Pt(11)
    run_sub.font.bold = True
    run_sub.font.color.rgb = RGBColor(71, 85, 105)
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # Athlete Scorecard Table
    t_profile = doc.add_table(rows=2, cols=4)
    t_profile.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = ["TARGET WEIGHT", "TARGET BODY FAT", "TARGET WAIST", "DAILY CALORIES & PROTEIN"]
    values = ["69.0 – 70.0 kg", "9.5% – 10.0%", "29.5\" – 30.0\"", "2,250–2,350 kcal | 145–155g P"]
    
    for i, h in enumerate(headers):
        c = t_profile.cell(0, i)
        set_cell_bg(c, "E0F2FE")
        p = c.paragraphs[0]
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(3, 105, 161)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        c2 = t_profile.cell(1, i)
        set_cell_bg(c2, "F8FAFC")
        p2 = c2.paragraphs[0]
        r2 = p2.add_run(values[i])
        r2.font.bold = True
        r2.font.size = Pt(10)
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 5-Day Workout Section
    for d in DAYS_DATA:
        h_day = doc.add_paragraph()
        r_day = h_day.add_run(f"{d['title']}: {d['subtitle']}")
        r_day.font.name = "Calibri"
        r_day.font.size = Pt(14)
        r_day.font.bold = True
        r_day.font.color.rgb = RGBColor(15, 23, 42)
        
        p_focus = doc.add_paragraph()
        r_focus = p_focus.add_run(f"Session Focus: {d.get('focus', d['subtitle'])} (Duration: {d['duration']})")
        r_focus.font.size = Pt(10)
        r_focus.font.italic = True
        r_focus.font.color.rgb = RGBColor(100, 116, 139)

        # Individual Exercise Cards with Embedded Visual GIF Demonstration
        for ex in d["exercises"]:
            card = doc.add_table(rows=1, cols=2)
            card.alignment = WD_TABLE_ALIGNMENT.CENTER
            c_img, c_info = card.rows[0].cells
            c_img.width = Inches(2.3)
            c_info.width = Inches(4.7)
            set_cell_bg(c_img, "F8FAFC")
            set_cell_bg(c_info, "FFFFFF")

            # Left Cell: GIF Image Demonstration
            gif_file = os.path.join(BASE_DIR, "gifs", ex["gif"])
            p_img = c_img.paragraphs[0]
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if os.path.exists(gif_file):
                try:
                    p_img.add_run().add_picture(gif_file, width=Inches(2.1))
                except Exception as e:
                    p_img.add_run(f"[GIF Demonstration: {ex['gif']}]")
            else:
                p_img.add_run(f"[GIF: {ex['gif']}]")
            
            p_b = c_img.add_paragraph()
            p_b.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_b.paragraph_format.space_before = Pt(2)
            p_b.paragraph_format.space_after = Pt(2)
            r_b = p_b.add_run(f"Exercise #{ex['num']} Visual Movement")
            r_b.font.size = Pt(8)
            r_b.font.bold = True
            r_b.font.color.rgb = RGBColor(100, 116, 139)

            # Right Cell: Exercise Coaching Information
            p0 = c_info.paragraphs[0]
            p0.paragraph_format.space_before = Pt(0)
            p0.paragraph_format.space_after = Pt(2)
            r_num = p0.add_run(f"#{ex['num']}. ")
            r_num.font.bold = True
            r_num.font.size = Pt(11)
            r_num.font.color.rgb = RGBColor(14, 116, 144)
            r_name = p0.add_run(f"{ex['name']}")
            r_name.font.bold = True
            r_name.font.size = Pt(11)
            r_name.font.color.rgb = RGBColor(15, 23, 42)

            p1 = c_info.add_paragraph()
            p1.paragraph_format.space_before = Pt(0)
            p1.paragraph_format.space_after = Pt(2)
            r_m_label = p1.add_run("Targeted Muscle: ")
            r_m_label.font.bold = True
            r_m_label.font.size = Pt(9.5)
            r_m_val = p1.add_run(ex['muscle'])
            r_m_val.font.bold = True
            r_m_val.font.size = Pt(9.5)
            r_m_val.font.color.rgb = RGBColor(22, 101, 52)

            p2 = c_info.add_paragraph()
            p2.paragraph_format.space_before = Pt(0)
            p2.paragraph_format.space_after = Pt(2)
            r_v_label = p2.add_run("Prescription: ")
            r_v_label.font.bold = True
            r_v_label.font.size = Pt(9)
            r_v_val = p2.add_run(f"{ex['sets_reps']}  |  {ex['rest']}")
            r_v_val.font.bold = True
            r_v_val.font.size = Pt(9)
            r_v_val.font.color.rgb = RGBColor(30, 64, 175)

            p_vid = c_info.add_paragraph()
            p_vid.paragraph_format.space_before = Pt(0)
            p_vid.paragraph_format.space_after = Pt(4)
            r_vid_l = p_vid.add_run("🎥 Video Tutorial: ")
            r_vid_l.font.bold = True
            r_vid_l.font.size = Pt(8.5)
            r_vid_v = p_vid.add_run(f"{ex.get('video_creator', 'YouTube')} - {ex.get('video_url', '')}")
            r_vid_v.font.size = Pt(8.5)
            r_vid_v.font.color.rgb = RGBColor(14, 116, 144)
            r_vid_v.font.underline = True

            p3 = c_info.add_paragraph()
            p3.paragraph_format.space_before = Pt(0)
            p3.paragraph_format.space_after = Pt(2)
            r_act_l = p3.add_run("🎯 Target Focus: ")
            r_act_l.font.bold = True
            r_act_l.font.size = Pt(8.5)
            r_act_v = p3.add_run(ex['activation'])
            r_act_v.font.size = Pt(8.5)

            p4 = c_info.add_paragraph()
            p4.paragraph_format.space_before = Pt(0)
            p4.paragraph_format.space_after = Pt(2)
            r_set_l = p4.add_run("📐 Stance & Setup: ")
            r_set_l.font.bold = True
            r_set_l.font.size = Pt(8.5)
            r_set_v = p4.add_run(ex['setup'])
            r_set_v.font.size = Pt(8.5)

            p5 = c_info.add_paragraph()
            p5.paragraph_format.space_before = Pt(0)
            p5.paragraph_format.space_after = Pt(2)
            r_form_l = p5.add_run("💡 Execution Cue: ")
            r_form_l.font.bold = True
            r_form_l.font.size = Pt(8.5)
            r_form_v = p5.add_run(ex['form'])
            r_form_v.font.size = Pt(8.5)

            p6 = c_info.add_paragraph()
            p6.paragraph_format.space_before = Pt(0)
            p6.paragraph_format.space_after = Pt(2)
            r_mis_l = p6.add_run("⚠️ Avoid Mistakes: ")
            r_mis_l.font.bold = True
            r_mis_l.font.size = Pt(8.5)
            r_mis_l.font.color.rgb = RGBColor(185, 28, 28)
            r_mis_v = p6.add_run(ex['mistakes'])
            r_mis_v.font.size = Pt(8.5)
            r_mis_v.font.color.rgb = RGBColor(185, 28, 28)

            p7 = c_info.add_paragraph()
            p7.paragraph_format.space_before = Pt(0)
            p7.paragraph_format.space_after = Pt(4)
            r_prg_l = p7.add_run("📈 Progression: ")
            r_prg_l.font.bold = True
            r_prg_l.font.size = Pt(8.5)
            r_prg_l.font.color.rgb = RGBColor(4, 120, 87)
            r_prg_v = p7.add_run(ex['progression'])
            r_prg_v.font.size = Pt(8.5)
            r_prg_v.font.color.rgb = RGBColor(4, 120, 87)

            p_sp = doc.add_paragraph()
            p_sp.paragraph_format.space_before = Pt(0)
            p_sp.paragraph_format.space_after = Pt(6)

        doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # Weekly Schedule
    h_sched = doc.add_paragraph()
    r_sched = h_sched.add_run("Weekly Schedule Alignment & Saturday Cricket Protocol")
    r_sched.font.size = Pt(14)
    r_sched.font.bold = True
    r_sched.font.color.rgb = RGBColor(15, 23, 42)

    p_sched_text = doc.add_paragraph()
    p_sched_text.add_run("Thursday Athletic Lower B is strategically positioned 48 hours before Saturday turf cricket. This provides full posterior chain recovery, protecting hamstrings and calves from match-day tears and cramps.")

    # Nutrition & 8 Staples
    h_nut = doc.add_paragraph()
    r_nut = h_nut.add_run("Daily Nutrition Targets & The 8 Non-Negotiable Staples")
    r_nut.font.size = Pt(14)
    r_nut.font.bold = True
    r_nut.font.color.rgb = RGBColor(15, 23, 42)

    staples_list = [
        "1. Up to 5 Whole Eggs / day (Never separate egg whites alone; provides bioavailable protein & choline).",
        "2. Up to 500ml Toned Milk / day (Split between morning oats and evening tea/coffee).",
        "3. 50g Pottukadalai (Roasted Chana - 4:30 PM crunch providing ~11g protein & 5g fiber).",
        "4. 15g Roasted Peanuts + 15g Raw Cashews (Zinc, magnesium, and essential healthy fats).",
        "5. 1–2 tsp Pure Honey (15g - morning oats sweetener or fast pre-gym energy).",
        "6. 50g Rolled / Instant Oats (Complex low-GI carbs for all-day steady energy).",
        "7. 45g Dry-weighed Soya Chunks (Weighed raw on scale before boiling; ~23.4g protein on non-chicken days).",
        "8. Chicken Routine (Wed & Sun): Single morning batch of home chicken gravy. No whey isolate powder needed!"
    ]
    for st in staples_list:
        p = doc.add_paragraph(style='List Bullet')
        r = p.add_run(st)
        r.font.size = Pt(9.5)

    # Morning APT Protocol
    h_apt = doc.add_paragraph()
    r_apt = h_apt.add_run("Daily 5-Minute Morning Posture & APT Reset Protocol")
    r_apt.font.size = Pt(14)
    r_apt.font.bold = True
    r_apt.font.color.rgb = RGBColor(15, 23, 42)

    apt_items = [
        "1. Stomach Vacuums: 4 sets × 20s hold (Strengthens transverse abdominis, flattens lower belly protrusion) — Video: https://www.youtube.com/watch?v=b5GsE6ZHn_E",
        "2. Kneeling Hip Flexor Stretch with Posterior Pelvic Tilt: 2 sets × 45–60s per leg (Lengthens tight psoas from WFH chair sitting) — Video: https://www.youtube.com/watch?v=gqoPYLUgP48",
        "3. Glute Squeezes & Posterior Pelvic Tilt Tucks: 2 sets × 10 reps (Activates glutes before sitting at IT desk) — Video: https://www.youtube.com/watch?v=mkbNJKiWNlI"
    ]
    for item in apt_items:
        p = doc.add_paragraph(style='List Bullet')
        r = p.add_run(item)
        r.font.size = Pt(9.5)

    doc.save(docx_path)
    print("DOCX generation complete!")

def main():
    html_out = os.path.join(BASE_DIR, "Master_Transformation_Guide.html")
    pdf_out = os.path.join(BASE_DIR, "Master_Transformation_Guide.pdf")
    docx_out = os.path.join(BASE_DIR, "Master_Transformation_Guide.docx")
    brain_html = os.path.join(BRAIN_DIR, "master_transformation_guide.html")

    print("Step 1: Writing Master_Transformation_Guide.html and index.html...")
    html_content = generate_html()
    with open(html_out, "w", encoding="utf-8") as f:
        f.write(html_content)
    index_out = os.path.join(BASE_DIR, "index.html")
    with open(index_out, "w", encoding="utf-8") as f:
        f.write(html_content)
    if os.path.exists(os.path.dirname(brain_html)):
        with open(brain_html, "w", encoding="utf-8") as f:
            f.write(html_content)
    curr_brain = r"C:\Users\SreeramThangavel\.gemini\antigravity\brain\bc1907da-f8a5-4573-b07f-fc5694e83446\master_transformation_guide.html"
    if os.path.exists(os.path.dirname(curr_brain)):
        with open(curr_brain, "w", encoding="utf-8") as f:
            f.write(html_content)
    print("HTML files written successfully.")

    print("Step 2: Compiling PDF via Playwright...")
    build_pdf(html_out, pdf_out)

    print("Step 3: Compiling formatted Word Document (.docx)...")
    build_docx(docx_out)

    print("\nALL MASTER DOCUMENTS SUCCESSFULLY GENERATED!")
    print(f"1. HTML: {html_out}")
    print(f"2. PDF:  {pdf_out}")
    print(f"3. DOCX: {docx_out}")
    print(f"4. Brain Artifact: {brain_html}")

if __name__ == "__main__":
    main()
