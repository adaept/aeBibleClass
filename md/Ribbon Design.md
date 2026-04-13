# Radiant Word Bible&#8482; — Navigation Ribbon

A custom toolbar for the Radiant Word Bible&#8482; Study document, built into Microsoft Word.
Navigate to any book, chapter, or verse in a few keystrokes or clicks.

---

## What it does

The Radiant Word Bible is a full-length Study Bible document — all 66 books, 1,189
chapters, and 31,102 verses — open in Microsoft Word. The navigation ribbon is a
custom toolbar that appears as its own tab, **Radiant Word Bible**, in the Word ribbon.

It gives you three ways to move through the text:

- Type or select the **book** you want
- Type or select the **chapter**
- Type or select the **verse**

Each step focuses the document at exactly that location.

---

## The navigation toolbar

```
◀  Genesis          ▼  ▶       ← Book row
◀  1                ▼  ▶       ← Chapter row
◀  1                ▼  ▶       ← Verse row

[New Search]    [About]
```

Each row has three controls: a **Previous** button, a **selector** (the text field
with a dropdown), and a **Next** button.

The rows unlock progressively:

| What you have done | What is available |
|--------------------|-------------------|
| Nothing yet | Book row only |
| Selected a book | Book row + Chapter row |
| Selected a chapter | Chapter row + Verse row |
| Selected a verse | All rows |

This keeps the choices simple. You cannot jump to a verse before you have confirmed
which book and chapter you are in.

---

## Selecting a book

Click the Book selector, or press **Alt, R W** to focus it directly from the keyboard.

You can:

- **Type a name** — full or abbreviated: `Genesis`, `Gen`, `Gn`, `1 Cor`, `1cor`,
  `Jn`, `Rev` all work. Capitalisation does not matter.
- **Open the dropdown** — all 66 books are listed, divided into Old Testament and
  New Testament.

Once the book is confirmed, the Chapter row becomes available and the document
moves to the first page of that book.

### Previous / Next Book

The **◀** and **▶** buttons beside the Book selector step one book backward or
forward. At Genesis, pressing ◀ does nothing. At Revelation, pressing ▶ does
nothing.

---

## Selecting a chapter

After a book is selected, the Chapter selector becomes active.

Type the chapter number, or open the dropdown to see all chapters in the current
book. The list reflects the actual chapter count for that book — Genesis offers
1–50, Jude offers only 1.

The document moves to the heading for that chapter and the Verse row becomes
available.

### Previous / Next Chapter

**◀** and **▶** step through chapters within the current book.

---

## Selecting a verse

After a chapter is selected, the Verse selector becomes active.

Type the verse number, or choose from the dropdown. The document moves directly to
that verse.

### Previous / Next Verse

**◀** and **▶** step through verses within the current chapter.

---

## Keyboard navigation

The ribbon is fully keyboard-navigable — no mouse required.

### Tab between fields

After typing in a selector, press **Tab** to confirm and move to the next selector:

```
Book field  →  Tab  →  Chapter field  →  Tab  →  Verse field  →  Tab  →  New Search
```

### Alt shortcuts (KeyTips)

Press **Alt** to activate the ribbon. Short letter badges appear on every control.
Press the letter shown to activate that control immediately, from anywhere in the
document.

| Key sequence | Action |
|--------------|--------|
| Alt, R W | Focus the Radiant Word Bible tab |
| Alt, R W, B | Focus the Book selector |
| Alt, R W, C | Focus the Chapter selector |
| Alt, R W, V | Focus the Verse selector |
| Alt, R W, S | New Search |
| Alt, R W, A | About |
| Alt, R W, [ | Previous Book |
| Alt, R W, ] | Next Book |
| Alt, R W, , | Previous Chapter |
| Alt, R W, . | Next Chapter |
| Alt, R W, < | Previous Verse |
| Alt, R W, > | Next Verse |

---

## New Search

The **New Search** button (or **Alt, R W, S**) clears all three selectors and
returns the toolbar to its starting state — Book row active, Chapter and Verse
rows inactive.

Use it whenever you want to navigate to a completely different passage. You do not
need to erase the selectors manually; New Search resets everything at once.

---

## Example: going to John 3:16

1. Press **Alt, R W, B** — Book selector is focused
2. Type `Jn` — the book resolves to John
3. Press **Tab** — Chapter selector is focused
4. Type `3`
5. Press **Tab** — Verse selector is focused
6. Type `16`
7. Press **Tab** — the document is at John 3:16

Or with the mouse: click the Book selector, type `Jn`, click the Chapter selector,
type `3`, click the Verse selector, type `16`.

---

## Example: reading through the Psalms

1. Navigate to Psalm 1:1 (as above)
2. Press **Alt, R W, .** (Next Chapter) to step to Psalm 2
3. Continue pressing **.** to read chapter by chapter
4. Press **Alt, R W, ,** (Previous Chapter) to go back

---

## System requirements

- Microsoft Word 365 (Windows)
- Macro-enabled document format (`.docm`)
- Macros must be enabled in Word's Trust Center for the document

---

## About

The Radiant Word Bible&#8482; navigation ribbon is part of the **adaept** Study Bible
project. The goal is a full-featured, keyboard-accessible Bible study environment
inside the familiar Word interface — no separate app required.

---

## Legal notices

Radiant Word Bible and RWB are trademarks of adaept. All rights reserved.

Microsoft and Microsoft Word are registered trademarks of Microsoft Corporation.
adaept is not affiliated with, endorsed by, or sponsored by Microsoft Corporation.
