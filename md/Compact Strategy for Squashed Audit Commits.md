
# ✅ Compact Strategy for Squashed Audit Commits

This page that should render cleanly in GitHub Markdown, VS Code, and static guide engines like MkDocs or Docusaurus:

---

## 📝 Example Commit Message Template

```text
Audit updates #299–#305
```

- Concise, readable, grep-friendly  
- Sufficient context for private or single-user workflows  
- Use task range only when applicable  

---

## 🛠️ Context: Demo Simplest Implementation Using GitHub Desktop + Git CLI

Since GitHub Desktop doesn't support squash merges directly, use the following workflow:

### 🔧 Step-by-Step

1. **Open the terminal for your repo**  
   GitHub Desktop: `Repository > Open in Terminal`

2. **Run this to squash the last N commits**  
   Replace `N` with the desired number:

   ```bash
   git reset --soft HEAD~N
   git commit -m "Audit updates #299–#305"
   ```

3. **Push the new single commit using GitHub Desktop**

---

## 📌 Example

You made 4 audit commits locally:

```bash
git reset --soft HEAD~4
git commit -m "Audit updates #312–#315"
```

Result: One tidy upstream commit after push.

---

## 🧩 Fixing Issues #289, #290, #291

### 🌿 1. Create a New Branch

Use GitHub Desktop:

```text
Branch > New Branch
```

- Name: `fix-audit-issues-289-291`  
- Click `Create Branch`

---

### 🧪 2. Implement Fixes with Separate Commits

For each issue:

1. Make edits for the fix  
2. In GitHub Desktop: `Changes` tab  
3. Commit message:  

   ```text
   Fix issue #289: [brief description]
   ```

4. Commit to `fix-audit-issues-289-291`

Repeat for:

- Fix for #290  
- Fix for #291  

---

### 🔀 3. Squash the Audit Commits

Open the terminal:

```bash
git reset --soft HEAD~3
git commit -m "Audit updates #289–#291"
```

---

### 🚀 4. Push Changes

In GitHub Desktop:

```text
Branch > Push Origin
```

---

### 🔁 5. Create a Pull Request

GitHub Desktop:

```text
Branch > Create Pull Request
```

- Title and description should link to issues #289, #290, and #291  
- Click `Create Pull Request`

---

## 🧩 Context: Active Branch `Sq274`

This branch carries the FIXED entry for #274, visible in the changelog block. There are other commits also FIXED for different tasks.

The final result can be seen here: <https://github.com/adaept/aeBibleClass/commit/47faa142c479a485167755cef65eb87290399504>

---

### 🔧 Squash Merge Workflow (Audit-Friendly)

#### 1. **Working on Sq274 in GitHub Desktop**

This branch will have a number of commits until it is ready to be processed for upstream audit log creation.

Current Status of Change Log:

``` Text
' FIXED - #302 - Update PrintCompactSectionLayoutInfo to output in rpt folder, move to basTESTaeBibleTools and add doc header
' FIXED - #301 - 999 AppendToFile should be "SKIPPED" [bug]
'Sq ' FIXED - #274 - Fix output path so 'Style Usage Distribution.txt' goes to rpt folder, add code header [bug] [doc]
```

The line with `'Sq` indicates the start of tasks to be squashed and merged.

#### 1.1 **Review and Consolidate Commits**

Before pushing squashed changes upstream, ensure the following:

- Review all commits in the branch to confirm they align with the intended scope of the task.  
- Consolidate related commits into a single, meaningful commit message that reflects the purpose of the changes.  
- Use GitHub Desktop to squash commits via the Pull Request workflow (manual squash not required).

#### 2. **Push Squashed Changes Upstream**

### 🔍 Step-by-Step: Squash for Audit Readiness

- Open the branch (`Sq274`) in GitHub Desktop and navigate to the **History** tab.  
- Review each commit to ensure it relates to the same task or changelog block.  
- Let the PR process handle the squash automatically—no manual squash needed.  
- In the **PR description**, write a concise summary of the consolidated change, using language that aligns with the audit log style.

🪶 Example:

**FIXED #274, #301, #302, 304** – Summary

**Extended Message Box:**  
FIXED - #304 - Add task type [wip] - it will prepend the task commits until replaced by FIXED
FIXED - #302 - Update PrintCompactSectionLayoutInfo to output in rpt folder, move to basTESTaeBibleTools and add doc header  
FIXED - #301 - 999 AppendToFile should be "SKIPPED" [bug]  
FIXED - #274 - Fix output path so 'Style Usage Distribution.txt' goes to rpt folder, add code header [bug] [doc]

#### 3. **Pull Request to Main**

- Create a PR from your local branch `Sq274` pushed to the remote repository on GitHub (`origin/Sq274`) targeting the remote `main` branch.  
- The PR compares the remote `Sq274` branch against the remote `main`, with squash handled server-side.  
- Carefully draft the title and description to reflect the consolidated task-level updates.

#### 4. **Merge with Squash Confirmed**

- Use "Squash and merge" in the PR UI to retain your audit-ready message.  
- Final commit in `main` should reflect the full summary and task identifiers for clarity.

---

### 🧭 Where to Find It

- After pushing your local branch (`Sq274`) to GitHub and opening a PR targeting `main`, scroll down to the **bottom of the PR page**.
- If there are no conflicts and all required checks pass (e.g., CI/CD, review approvals), GitHub displays a green **“Merge pull request”** button.
- Just above or beside that, there’s a dropdown arrow with merge strategy options.

Click the dropdown—then select:

> ✅ **Squash and merge**  
> 🔁 Consolidates all commits into one before merging to `main`, using your drafted title and description.

---

### 🧪 What You Control

- You’ll be prompted to **edit the final commit message** before confirming the merge.

This is your moment to paste the audit-friendly message from your changelog block.

---

#### 5. **Post-Merge Local Cleanup (GitHub Desktop Only)**

After completing the squash merge in GitHub, perform these post-merge steps using GitHub Desktop:

1. **Delete Local Branch `Sq274`**
   - Open **Branches** tab in GitHub Desktop.
   - Right-click on `Sq274` → **Delete Branch**.

2. **Sync Local `main` with Remote**
   - Switch to the `main` branch.
   - Click **Fetch Origin** → then **Pull Origin**.

3. **Verify Local History**
   - Go to the **History** tab.
   - Confirm latest commits match remote `main`.

4. **Optional: Housekeeping**
   - Review stale branches in **Branches** tab.
   - Delete unused ones if desired.
