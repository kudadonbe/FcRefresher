# Office 365 Productivity Training
## Hybrid Session - 29 December 2025

---
marp: true
theme: uncover
class: invert
paginate: true
backgroundImage: url('../../settings/kudadonbe_theme_black.png')
---

<!-- _class: lead -->
# **OFFICE 365 PRODUCTIVITY BOOST 2026**
### Smart Digital Tools for Efficient Court Work
#### Hybrid Training Session

**29 December 2025**
Kaiveni Hall, Henveyru & Microsoft Teams

---
<!-- _class: lead -->
# **Today's Presenter**

**Hussain Shareef**  
Computer Technician | ICT Department

📞 +960 3340507  
📧 hussain.shareef@familycourt.gov.mv  
🌐 www.familycourt.gov.mv

---
<!-- _class: lead -->
# **Hybrid Session Setup**

## **In-Person:**
📍 Kaiveni Hall, Henveyru  
🎤 Microphone & Projector  
📄 Printed Materials

## **Remote:**
💻 Microsoft Teams Meeting  
🔗 [Teams Link Here]  
📱 Meeting ID: [ID Here]

---
# **Session Objectives**

By the end of today, you will:

1. ✅ **Implement** proper file naming conventions
2. ✅ **Choose** correct storage (OneDrive vs SharePoint)
3. ✅ **Apply** XLOOKUP formulas in Excel
4. ✅ **Master** Word document cleanup
5. ✅ **Navigate** hybrid collaboration tools

---
# **Session Agenda**

| Time | Activity | For Remote Participants |
|------|----------|-------------------------|
| 8:30 | Welcome & Tech Check | Test audio/video |
| 8:35 | File Naming Standards | Follow shared screen |
| 8:43 | Storage Strategy | Interactive poll |
| 8:50 | Excel: XLOOKUP | Download practice file |
| 8:58 | Word Cleanup | Live demo |
| 9:05 | Quick Quiz | Teams Forms |
| 9:10 | Activities | Breakout rooms |
| 9:35 | Q&A | Use chat feature |
| 9:40 | Closing | Links in chat |

---
<!-- _class: lead -->
# **The File Naming Crisis**

## Can You Find This File?

📁 `Hassan_v_Mohamed_final_FINAL_revised_2_actuallyfinal.docx`

---
# **Common Problems**

❌ **260-character** Windows path limit  
❌ Difficult to back up/recover  
❌ Confusing version control  
❌ Hard to search efficiently  
❌ Archive nightmares

---
<!-- _class: lead -->
# **Raise Your Hand If...**

☐ Spent >5 minutes searching for a file this week  
☐ Received "file name too long" error  
☐ Shared wrong version accidentally  
☐ Created duplicate files unknowingly

---
# **Family Court Naming Standard**
### *Effective 1 January 2026*

```
YYYY-MM-DD_CaseType-ID_Document_Initials
```

## **Examples:**
✅ `2025-12-29_FC-2025-0456_Motion_HSR.docx`  
✅ `2025-11-15_JD-2025-0789_Order_JAS.docx`  
✅ `2025-10-20_CA-2025-1234_Report_MAA.docx`

---
# **Benefits of Standard Naming**

✨ **Chronological sorting** automatically  
✨ **Easy identification** at a glance  
✨ **Consistent** across entire court  
✨ **Future-proof** for archiving  
✨ **Search-friendly** for all staff

---
<!-- _class: lead -->
# **Case Type Codes**

**FC** = Family Case  
**JD** = Juvenile Division  
**CA** = Child Affairs  
**AP** = Appeal Case  
**MO** = Maintenance Order

---
# **Metadata: The Hidden Power**

## **How to Add Metadata:**
1. Right-click file → **Properties**
2. Click **Details** tab
3. Add/Edit fields:
   - Authors
   - Tags (keywords)
   - Comments
   - Subject

---
# **Court-Specific Metadata Tags**

🔑 **Case Number**  
👨‍⚖️ **Judge Name**  
📅 **Hearing Date**  
⚠️ **Urgency Level**  
📂 **Case Category**  
👥 **Parties Involved**

---
<!-- _class: lead -->
# **Search Magic**

## Find files by **CONTENT**, not just filename!

Metadata enables powerful search:
- "Show me all urgent Family Cases"
- "Find Judge Ibrahim's December orders"
- "Locate maintenance orders from 2025"

---
# **Storage Strategy**
## Where Should This File Go?

---
# **🟦 OneDrive**
### Your Personal Work Locker

✅ Draft documents  
✅ Personal notes  
✅ Work-in-progress  
✅ Only **YOU** control access  
✅ Auto-saves every 2 seconds  
✅ Accessible anywhere

---
# **🟪 SharePoint**
### Team Shared Library

✅ Court templates  
✅ Finalized forms  
✅ Team documents  
✅ Approved procedures  
✅ **Everyone** in team can access  
✅ Version history for all

---
<!-- _class: lead -->
# **Golden Rule**

# **Drafting → OneDrive**
# **Final/Team → SharePoint**

---
# **Quick Demo: Moving Files**

1. **Find** file in OneDrive
2. **Drag & drop** to SharePoint
3. **Verify** team access
4. **Delete** from OneDrive (optional)

---
<!-- _class: lead -->
# **Excel Superpower: XLOOKUP**

## Your New Best Friend for Case Management

---
# **Perfect For:**

📊 **Year-end case statistics**  
🔍 **Pulling case details**  
📈 **Generating reports**  
📎 **Finding related documents**  
🔄 **Data validation and cleanup**

---
# **Simple Formula**

```excel
=XLOOKUP(What, Where, Return, "Not Found")
```

## **Court Example:**
```excel
=XLOOKUP(A2, '2025_Cases'!A:A, '2025_Cases'!C:C, "Check Archive")
```

---
# **Live Demo Scenario**

**Task:** Find case officer by Case ID

**Data:** 
- Column A: Case Numbers (FC-2025-001, FC-2025-002...)
- Column B: Case Officers (Aminath, Fathimath, Ibrahim, Mohamed, Aishath...)

**Formula:**
```excel
=XLOOKUP("FC-2025-0456", A:A, B:B, "Officer Not Found")
```

---
<!-- _class: lead -->
# **XLOOKUP vs VLOOKUP**

## Why XLOOKUP Wins for Court Work

---
| **VLOOKUP Limitations** | **XLOOKUP Advantages** |
|-------------------------|------------------------|
| ❌ Can only look right | ✅ Looks left/right/up/down |
| ❌ Breaks if columns change | ✅ Handles column changes |
| ❌ Requires exact column count | ✅ Simpler syntax |
| ❌ Slower with large data | ✅ Faster processing |
| ❌ No built-in error handling | ✅ Built-in "if not found" |
| ❌ Hard to maintain | ✅ Easy to update |

---
<!-- _class: lead -->
# **Your 2026 Resolution**

# **Replace ALL VLOOKUPs with XLOOKUP!**

---
# **Word Document Cleanup**
## Professional Documents in 60 Seconds

---
# **The Magic '¶' Button**

## **Keyboard Shortcut:**
`Ctrl + Shift + 8`

## **Or Click:**
Home tab → Paragraph group → **¶**

---
# **What It Reveals:**

🔍 **Extra spaces** (dots between words)  
🔍 **Tab characters** (→ arrows)  
🔍 **Paragraph marks** (¶ symbols)  
🔍 **Hidden formatting**  
🔍 **Manual line breaks**

---
# **Common Formatting Issues**

❌ Inconsistent spacing  
❌ Manual numbering (1., 2., 3.)  
❌ Mixed fonts and sizes  
❌ Poor alignment  
❌ Missing styles  
❌ Random bold/italic

---
# **60-Second Cleanup Process**

1. **Show** formatting (`Ctrl+Shift+8`)
2. **Remove** extra spaces
3. **Apply** "Court Style" template
4. **Update** headers/footers
5. **Add** metadata
6. **Save** as new 2026 template

---
<!-- _class: lead -->
# **Interactive Quiz Time!**

## Test Your Knowledge

---
# **Question 1**
**Best filename for 2026 court documents?**

A) `finaldocument.docx`  
B) `document1_new_final.docx`  
C) `2026-01-15_FC2026-001_Motion_HSR.docx`  
D) `motion_for_hearing.docx`

---
# **Question 2**
**Where should final court forms be saved?**

A) Desktop  
B) OneDrive  
C) SharePoint  
D) Email attachments

---
# **Question 3**
**Quick key to show Word formatting marks?**

A) `Ctrl+F`  
B) `Ctrl+Shift+8`  
C) `Windows+L`  
D) `Alt+F4`

---
# **Answers:**

1. **C** `2026-01-15_FC2026-001_Motion_HSR.docx`
2. **C** SharePoint
3. **B** `Ctrl+Shift+8`

---
<!-- _class: lead -->
# **Activity Time!**

## Hands-On Practice Session

---
# **Activity 1: File Naming Challenge**
### *5 minutes*

**Task:** Rename 5 messy files using court standard

**Format:** `YYYY-MM-DD_CaseType-ID_Document_Initials`

**Prize:** Premium USB drive for fastest correct solution

---
# **Activity 2: Storage Sorting Game**
### *7 minutes*

**Setup:** 
- Blue basket = OneDrive
- Purple basket = SharePoint
- 10 file scenario cards

**Task:** Decide where each file should go

**Prize:** Wireless mouse

---
# **Activity 3: Document Cleanup Race**
### *3 minutes*

**Task:** Fix formatting in a messy document

**Requirements:**
1. Show formatting marks
2. Remove extra spaces
3. Apply Court Style
4. Save with correct filename

**Prize:** Ergonomic keyboard

---
# **Activity 4: Excel Formula Builder**
### *8 minutes*

**Task:** Create XLOOKUP formula for case lookup

**Formula:**
```excel
=XLOOKUP(lookup_value, lookup_array, return_array, "not_found")
```

**Prize:** Excel reference book

---
# **For Remote Participants:**

## **Virtual Activities via Teams**

1. **Breakout rooms** for each activity
2. **Shared screen** for demonstrations
3. **Chat support** from facilitator
4. **Digital prizes** for winners

---
<!-- _class: lead -->
# **2026 Support Plan**

## Continuing Your Learning Journey

---
# **Regular Support Schedule**

📅 **Tech Clinic:** Every Tuesday, 10:00-12:00  
📧 **Email Support:** Response within 24 hours  
📞 **Urgent Issues:** Call x5050 or +960 3340507  
💬 **Teams Chat:** Available during work hours

---
# **Available Resources**

📁 **Quick Guides:** `SharePoint/Training/`  
🎬 **Video Tutorials:** Court Intranet  
📋 **Templates:** `SharePoint/Templates/`  
📖 **Reference Cards:** ICT Department

---
# **Upcoming January Sessions**

1. **Automating Court Reports with Power Tools**
2. **Teams Collaboration & Meetings Mastery**
3. **Data Security & Privacy Basics**
4. **Advanced Excel for Case Management**

---
<!-- _class: lead -->
# **Q&A Session**

## Your Questions, Our Answers

---
# **How to Ask Questions:**

**In-Person:** Raise your hand  
**Remote:** 
- Use **Raise Hand** feature
- Type in **Chat**
- Unmute and speak

**After Session:** Email your questions

---
# **Contact Information**

**Hussain Shareef**  
Computer Technician | ICT Department

📞 **Phone:** +960 3340507  
📧 **Email:** hussain.shareef@familycourt.gov.mv  
🏢 **Office:** Justice Building, 20212, Malé  
🌐 **Website:** www.familycourt.gov.mv

---
# **Session Resources**

🔗 **Teams Recording:** Available in Chat  
📄 **Handouts:** `SharePoint/Training/29Dec2025/`  
📊 **Practice Files:** Download from Chat  
📝 **Feedback Form:** Link in Chat

---
<!-- _class: lead -->
# **Thank You!**

## Let's Make 2026 Our Most Productive Year Yet!

**Remember:** Start implementing these tips **tomorrow!**

---
# **Quick Reference Summary**

1. **File Names:** `YYYY-MM-DD_CaseType-ID_Doc_Initials`
2. **Storage:** Drafts→OneDrive, Final→SharePoint
3. **Excel:** Use XLOOKUP, not VLOOKUP
4. **Word:** `Ctrl+Shift+8` for formatting cleanup
5. **Support:** Tuesdays 10-12, email, or call

---
# **Feedback Matters!**

## Please Complete Our Short Survey:

📋 **Link in Teams Chat**  
📋 **QR Code on Screen**  
📋 **Email follow-up**

Your feedback shapes future training sessions!

---
<!-- _class: lead -->
# **One Last Thing...**

## Your Digital Resolution for 2026:

**"I will implement at least ONE thing I learned today, starting tomorrow."**

Which one will it be?

---
# **Safe Travels & Happy New Year!**

🎉 **Wishing You a Productive 2026!** 🎉

**Next Session:** January 2026  
**Topic:** Automating Court Reports

---
# **Credits & Acknowledgements**

**Presenter:** Hussain Shareef  
**Venue:** Kaiveni Hall, Henveyru  
**IT Support:** Family Court ICT Department  
**Materials:** Family Court Training Resources

**© 2025 Family Court, Republic of Maldives**

---
<!-- _class: lead -->
## **End of Presentation**

# **Questions?**