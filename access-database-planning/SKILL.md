# Access Database Planning Skill

Plan, structure, and document Microsoft Access database features using a two-phase approach: Product Requirements Documents (PRD) and task-driven implementation.

## When to Use This Skill

Use this skill when you need to:
- **Plan a new Access feature** from a user request or idea
- **Break down complex requirements** into manageable tasks
- **Document design decisions** before implementation begins
- **Guide a junior developer** through a feature with clear, actionable steps
- **Scope and clarify ambiguous requests** before coding starts

## How It Works

This skill uses a two-phase planning workflow:

### Phase 1: PRD (Product Requirements Document)
Create a detailed specification document that clarifies **what** needs to be built and **why**.

### Phase 2: Task List
Break the PRD or requirements into step-by-step tasks that guide **how** to implement the feature.

---

## Phase 1: Creating a PRD

### When to Create a PRD

Always create a PRD when:
- The user's request is vague or open-ended
- The feature is complex or touches multiple components
- You need stakeholder/user clarity before implementation
- The feature has unclear scope or success criteria

### PRD Process

1. **Receive Initial Request:** User provides a brief feature description or request
2. **Ask Clarifying Questions:** Ask only 3-5 essential questions focused on:
   - **Problem/Goal:** What problem does this solve?
   - **Core Functionality:** What can users do with this?
   - **Scope/Boundaries:** What should this NOT do?
   - **Success Criteria:** How do we measure success?
3. **Generate PRD:** Use the PRD Structure (see below) based on user answers
4. **Save:** Store as `prd-[feature-name].md` in the `/tasks` directory

### PRD Structure

The generated PRD must include these sections:

```markdown
# [Feature Name] - Product Requirements Document

## Introduction/Overview
Brief description of the feature and the problem it solves.

## Goals
List specific, measurable objectives for this feature.

## User Stories
Detail user narratives describing feature usage and benefits.
Example: "As a [user type], I want [capability] so that [benefit]"

## Functional Requirements
Numbered list of specific functionalities the feature must have.
- FR1: The system must allow users to [specific action]
- FR2: The system must validate [specific input]

## Non-Goals (Out of Scope)
Clearly state what this feature will NOT include.

## Design Considerations (Optional)
UI/UX requirements, mockups, relevant components, styling guidelines.

## Technical Considerations (Optional)
Known constraints, dependencies, suggested architecture.

## Success Metrics
How will success be measured? (e.g., "Reduce support tickets by 20%")

## Open Questions
Remaining ambiguities or areas needing further clarification.
```

### Clarifying Questions Format

When asking questions, use this format:

```
1. [Question text]
   A. [Option A]
   B. [Option B]
   C. [Option C]
   D. [Option D]

2. [Question text]
   A. [Option A]
   B. [Option B]
```

Users respond with selections like "1A, 2B, 3C" for easy reference.

---

## Phase 2: Generating Task Lists

### When to Create Task Lists

Create a task list when you have:
- An approved PRD
- Clear functional requirements
- A well-defined scope
- Need to guide implementation

### Task List Process

1. **Analyze Requirements:** Review PRD or requirements document
2. **Generate Parent Tasks:** Create 4-6 high-level tasks:
   - **Always include Task 0.0:** "Create feature branch" (unless user specifically opts out)
   - Example parent tasks: "Design Database Schema", "Build Form UI", "Implement Business Logic", "Write Tests", "Documentation"
3. **Present to User:** Show parent tasks and ask "Ready to generate sub-tasks? Respond with 'Go' to proceed"
4. **Generate Sub-Tasks:** Break each parent task into 2-5 actionable sub-tasks
5. **Identify Relevant Files:** List all files that will be created or modified
6. **Save:** Store as `tasks-[feature-name].md` in the `/tasks` directory

### Task List Structure

```markdown
# [Feature Name] - Implementation Tasks

## Relevant Files

- `path/to/file1.accdb` - Brief description of why relevant
- `path/to/module1.bas` - Description of module purpose
- `path/to/form1.frm` - Form code module
- `path/to/test/test_feature.ps1` - PowerShell tests for this feature

### Notes

- Store VBA modules alongside form/report files for version control
- PowerShell tests use the AccessPOSH module for COM automation
- Update this checklist after completing each sub-task (mark `- [ ]` → `- [x]`)

## Tasks

- [ ] 0.0 Create feature branch
  - [ ] 0.1 Create and checkout a new branch (e.g., `git checkout -b feature/[feature-name]`)

- [ ] 1.0 [Parent Task Title]
  - [ ] 1.1 [Sub-task description]
  - [ ] 1.2 [Sub-task description]

- [ ] 2.0 [Parent Task Title]
  - [ ] 2.1 [Sub-task description]
  - [ ] 2.2 [Sub-task description]
```

### Completing Tasks

As you complete each task:
1. Change `- [ ]` to `- [x]` in the markdown
2. Update after completing each sub-task (not just parent tasks)
3. Verify the work passes testing before marking complete

---

## Best Practices

### For PRDs
- **Be explicit:** Avoid jargon; assume a junior developer will read this
- **Focus on the "what" and "why":** Let developers figure out the "how"
- **Ask only essential questions:** Don't over-engineer clarification
- **Don't start implementing:** PRD is a planning artifact, not implementation

### For Task Lists
- **Actionable titles:** Each task should be concrete and achievable
- **Logical sequence:** Sub-tasks should flow naturally from the parent task
- **Clear dependencies:** Note if one task must complete before another
- **Testing integration:** Include testing tasks within parent tasks, not as an afterthought
- **File specificity:** Name the exact VBA modules, forms, or files being modified

### Access-Specific Guidance
- **Naming:** Always follow [vba-naming.instructions.md](../_instructions/vba-naming.instructions.md) when naming tables, fields, controls, modules, procedures
- **Reserved words:** Check [access-vba-reserved-words SKILL.md](../access-vba-reserved-words/SKILL.md) for naming collisions before implementation
- **PowerShell testing:** Use AccessPOSH module for test automation (COM interop)
- **Form/Control tasks:** Break down into separate tasks for data layer, UI, and logic

---

## Output Files

| Phase | Output File | Location |
|-------|------------|----------|
| PRD | `prd-[feature-name].md` | `/tasks/` |
| Tasks | `tasks-[feature-name].md` | `/tasks/` |

---

## Example Workflow

**User Request:** "I want to add a customer search feature to our database"

**Step 1: PRD Phase**
- Ask clarifying questions: "Should search be real-time or button-triggered?", "Which fields?", "Should results be paginated?"
- Generate PRD with user answers
- Save as `prd-customer-search.md`

**Step 2: Task Phase**
- Analyze PRD requirements
- Generate parent tasks (Create feature branch, Design schema, Build search form, Implement search logic, Write tests)
- Wait for user "Go" confirmation
- Generate detailed sub-tasks for each parent
- Save as `tasks-customer-search.md`

**Step 3: Implementation**
- Follow task list sequentially
- Check off tasks as completed
- Run tests to verify before marking "done"
