تشغيل مخزن حجاج من السحابة (من غير جهازك)
==========================================
الفكرة: ملفات الإكسيل على Google Drive ← GitHub بيقراها كل ~10 دقايق ← الموقع على Cloudflare Pages.
لو الأرقام اتغيرت بس بينشر نسخة جديدة. ومفيش حاجة بتشتغل على جهازك.

1) Google Drive
   - اعمل فولدر على الدرايف وحط فيه ملفات الإكسيل (أي ملف .xlsx، أو Google Sheet).
   - افتح console.cloud.google.com ← مشروع جديد ← APIs & Services ← Library ← فعّل "Google Drive API".
   - IAM & Admin ← Service Accounts ← Create ← أي اسم ← بعد ما يتعمل: Keys ← Add key ← JSON (هينزل ملف). ده سري.
   - انسخ إيميل الـ service account (شكله xxx@xxx.iam.gserviceaccount.com) وشارك معاه الفولدر كـ "Viewer".
   - رقم الفولدر: من لينك الفولدر، الجزء اللي بعد folders/

2) Cloudflare
   - حساب مجاني ← Workers & Pages ← Create ← Pages ← Upload assets ← اسم المشروع (مثلًا makhzan) ← ارفع أي ملف مؤقت وسيب.
   - My Profile ← API Tokens ← Create Token ← Custom ← Account / Cloudflare Pages / Edit.
   - Account ID: من الصفحة الرئيسية للحساب.

3) GitHub
   - حساب مجاني ← New repository ← يفضل Public (الكود بس هو الظاهر، مفيش بيانات فيه، وكده التشغيل مجاني بلا حد).
   - ارفع محتويات الفولدر ده كله (بما فيه فولدر .github) في الـ main.
   - Settings ← Secrets and variables ← Actions ← New repository secret، وضيف 6:
       GOOGLE_SERVICE_ACCOUNT_JSON   = كل محتوى ملف الـ JSON
       DRIVE_FOLDER_ID               = رقم الفولدر
       CF_PROJECT                    = اسم مشروع كلاودفلير
       CLOUDFLARE_API_TOKEN          = التوكن
       CLOUDFLARE_ACCOUNT_ID         = الـ Account ID
       CF_URL                        = عنوان الموقع (مثلًا makhzan-ar8.pages.dev)
   - تبويب Actions ← sync-makhzan ← Run workflow. لما يخلص، الرابط: https://<اسم المشروع>.pages.dev

بعد كده: عدّل الإكسيل في الدرايف من أي جهاز ← خلال ~10-15 دقيقة الموقع يتحدث. ملف جديد تحطه في الفولدر بيبقى قسم لوحده تلقائي.
لو ملف مش مفهوم بيظهر تنبيه ⚠ في الصفحة الرئيسية. لو ظهر خطأ شوف سجل التشغيل في Actions.
ملاحظة: التحديث مش لحظي زي السيرفر المحلي، والصفحة بتتأكد من التحديث كل 5 ثواني.
