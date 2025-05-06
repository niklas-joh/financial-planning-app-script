# GitHub CLI Workflow Guide

This guide provides instructions on how to use GitHub CLI (`gh`) for managing issues, branches, commits, and more in this repository.

## 1. Managing GitHub Issues

### Creating Issues
```bash
# Create a new issue with title and body
gh issue create --title "Issue title" --body "Issue description"

# Create with labels and assignees
gh issue create --title "Fix dropdown bug" --body "The dropdown menu isn't working correctly" --label "bug,priority" --assignee "niklas-joh"
```

### Viewing Issues
```bash
# List open issues
gh issue list

# View a specific issue
gh issue view ISSUE_NUMBER
```

### Closing Issues
```bash
# Close an issue
gh issue close ISSUE_NUMBER

# Close with comment
gh issue close ISSUE_NUMBER --comment "Fixed in latest commit"
```

## 2. Branch Management

### Creating Branches
```bash
# Create and checkout a new branch
git checkout -b branch-name

# Create branch from an issue
gh issue develop ISSUE_NUMBER -b branch-name
```

### Listing Branches
```bash
# List all branches
git branch

# List remote branches
git branch -r
```

## 3. Committing Code

```bash
# Stage all changes
git add .

# Stage specific files
git add file1.js file2.js

# Commit with message
git commit -m "Fix: Resolve dropdown selection issue"
```

### Commit Message Best Practices

1. **Use a structured format**:
   ```
   <type>: <short summary>
   
   [optional body]
   
   [optional footer]
   ```

2. **Common types**:
   - `feat`: A new feature
   - `fix`: A bug fix
   - `docs`: Documentation changes
   - `style`: Changes that don't affect code meaning (formatting, etc.)
   - `refactor`: Code changes that neither fix a bug nor add a feature
   - `perf`: Performance improvements
   - `test`: Adding or correcting tests
   - `chore`: Changes to build process or auxiliary tools

3. **Keep the summary concise** (under 50 characters if possible)

4. **Use imperative mood** in the subject line ("Add feature" not "Added feature")

5. **Reference issues in the footer**:
   ```
   fix: Resolve dropdown selection issue
   
   The dropdown now maintains its state after form submission
   
   Fixes #42
   ```

## 4. Pushing Code

```bash
# Push to remote
git push origin branch-name

# Push and set upstream
git push -u origin branch-name
```

## 5. Pull Requests (for solving issues)

### Creating PRs
```bash
# Create PR for current branch
gh pr create --title "Fix dropdown functionality" --body "Resolves #42"

# Create with labels and reviewers
gh pr create --title "Fix dropdown functionality" --body "Resolves #42" --label "bugfix" --reviewer "colleague-username"
```

### Managing PRs
```bash
# List pull requests
gh pr list

# View a specific PR
gh pr view PR_NUMBER

# Check out a PR locally
gh pr checkout PR_NUMBER

# Merge a PR
gh pr merge PR_NUMBER
```

## Complete Workflow Example

```bash
# 1. Create an issue
gh issue create --title "Fix dropdown selection bug" --body "The dropdown menu doesn't maintain selection state"

# 2. Create a branch for the issue
gh issue develop 42 -b fix-dropdown-selection

# 3. Make your changes to the code
# ... edit files ...

# 4. Commit your changes
git add src/features/dropdowns.js
git commit -m "Fix: Maintain dropdown selection state"

# 5. Push your changes
git push -u origin fix-dropdown-selection

# 6. Create a pull request
gh pr create --title "Fix dropdown selection bug" --body "Resolves #42"

# 7. Merge the PR when approved
gh pr merge 7
```

## Best Practices for GitHub Workflow

1. **Always work in feature branches**, never directly on main/master
2. **Keep branches short-lived** and focused on a single issue or feature
3. **Pull the latest changes** from the main branch before creating a new branch
4. **Write descriptive commit messages** following the conventional commits format
5. **Reference issue numbers** in commits and PRs using keywords like "Fixes #42" or "Resolves #42"
6. **Request reviews** from team members on important changes
7. **Use labels** to categorize issues and PRs
8. **Delete branches** after they've been merged

This workflow integrates GitHub issues, branches, commits, and PRs to provide a complete solution for managing your development process through the command line.
