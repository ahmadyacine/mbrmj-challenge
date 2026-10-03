// Renders the report email from backend.gs with sample data -> email-preview.html
// Run:  node preview-email.js      (then open email-preview.html in a browser)
const fs = require("fs"), vm = require("vm");
const src = fs.readFileSync(__dirname + "/backend.gs", "utf8");
const ctx = {}; vm.createContext(ctx);
vm.runInContext(src.slice(src.indexOf("var EMAIL_LOGO"), src.indexOf("/* ---------------- setup + dashboard")), ctx);
const sample = {
  name: "Ahmad Yacine",
  report: {
    name: "Ahmad Yacine", score: "٤/٨",
    rows: [
      { ar: "تحليل المتطلبات", st: "blind", label: "نقطة عمياء" }, { ar: "التصحيح", st: "blind", label: "نقطة عمياء" },
      { ar: "الاختبار", st: "gap", label: "فجوة" }, { ar: "التصميم", st: "strong", label: "جاهز" }, { ar: "الكود الموجود", st: "strong", label: "جاهز" }
    ],
    focus: { ar: "تحليل المتطلبات", st: "blind", label: "نقطة عمياء", m: "كنت متأكد إنك فاهم المطلوب، بس فاتك شي أساسي. في الشغل، هذا يعني تبني الشي الغلط بثقة وتعيد أسبوع شغل.", n: "قبل أي مهمة، اكتب ٥ أسئلة للعميل وصنّف متطلباته (وظيفي / غير وظيفي)." },
    others: [
      { ar: "التصحيح", st: "blind", label: "نقطة عمياء", m: "كنت متأكد من إجابتك وطلعت غلط — وهذا أخطر نوع. في الشغل، يعني تسلم كود فيه Bug وأنت مطمّن.", n: "تدرب على قراءة الشروط سطر سطر، وجرب كل دالة بقيم حقيقية على الورق قبل ما تحكم." },
      { ar: "الاختبار", st: "gap", label: "فجوة", m: "تعرف إن الاختبار مهم، بس ما تعرف وش تختبر بالضبط.", n: "تعلّم الفرق بين Happy path و Edge case و Error case، وطبقها على ميزة وحدة." }
    ],
    strong: [{ ar: "التصميم", m: "تفكر بالتصميم قبل الكود، وتشوف الحالات اللي غيرك ما يشوفها." }, { ar: "الكود الموجود", m: "تعرف تدخل كود ما كتبته بدون ما تخاف أو تخرّب — وهذا ٩٩٪ من أول وظيفة." }],
    story: {
      situation: "لقيت Bug في نظام حجز يسمح بحجزين متداخلين على نفس الملعب.",
      did: "شغلت المشروع وكررت المشكلة، وتتبعت دالة التحقق. لقيت إنها أول ما تشوف حجز ملغي ترجع “متاح” وتوقف، بدون ما تشيك باقي الحجوزات. غيرت return True إلى continue، وأضفت Exclusion Constraint في قاعدة البيانات عشان الطلبات اللي تجي بنفس الثانية.",
      result: "كتبت ٣ اختبارات: واحد يعيد الخطأ نفسه ويثبت إنه انحل، وواحد يتأكد إن الأوقات الملغية تنحجز عادي، وواحد يتأكد إن الحجز اللي بعده مباشرة ينقبل.",
      tail: "هذي قصة من سؤال واحد. تخيّل قصتك بعد مشروع كامل."
    },
    plan: [
      ["اليوم ١", "اكتب ٥ أسئلة تسألها لأي عميل قبل ما تبدأ مهمة جديدة"], ["اليوم ٢", "خذ تطبيق تستخدمه يومياً واكتب ٥ متطلبات وظيفية له"], ["اليوم ٣", "اكتب ٣ متطلبات غير وظيفية لنفس التطبيق وصنّف كل شي"],
      ["اليوم ٤", "خذ دالة فيها شرط وجرّبها بقيم حقيقية على الورق"], ["اليوم ٥", "كرر Bug واكتب: القيم الداخلة، المتوقع، الفعلي"], ["اليوم ٦", "دوّر على المشكلة سطر سطر وسجّل كل فرضية جربتها"],
      ["اليوم ٧", "اكتب قصتك للمقابلة بنفس القالب عن أي مشروع سويته"]
    ],
    closing: "تقدر تتعلم هذي لحالك. بس في الشغل الحقيقي فيه أحد يراجع شغلك ويقولك وين غلطت — وهذا اللي تعطيك إياه دفعة مبرمج: مشروع كامل، مراجعة كود، نشر، وقصة حقيقية للمقابلة.",
    checklist: { url: "https://ahmadyacine.github.io/mbrmj-challenge/Mbrmj-Real-Task-Checklist.pdf", text: "حمّل قائمة “المهمة الحقيقية” PDF" },
    cta: { url: "https://mbrmj.ae", text: "سجّل في قائمة انتظار الدفعة القادمة" }
  }
};
let html = ctx.buildEmail_(sample);
html = html.replace(ctx.EMAIL_LOGO, "logo-white.png"); // local logo for the preview
fs.writeFileSync(__dirname + "/email-preview.html", html);
console.log("wrote email-preview.html");
