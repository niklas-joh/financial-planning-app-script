# Cline Instructions

This file contains instructions and best practices for working with this repository, particularly focusing on GitHub CLI workflows and git best practices.

## GitHub CLI Usage

For detailed GitHub CLI usage, refer to [github_cli_workflow.md](./github_cli_workflow.md) in this directory.

## Git Workflow Best Practices

### Branch Management

1. **Always work in feature branches**
   - Never commit directly to main/master
   - Create a new branch for each feature or bugfix

2. **Use descriptive branch names**
   - Format: `type/short-description`
   - Examples: `feature/add-dropdown-filter`, `fix/dropdown-selection-bug`

3. **Keep branches short-lived**
   - Merge or delete branches once the feature is complete
   - Regularly pull changes from main to avoid divergence

### Commit Practices

1. **Make atomic commits**
   - Each commit should represent a single logical change
   - This makes it easier to review, revert, or cherry-pick changes

2. **Follow conventional commit format**
   ```
   <type>: <short summary>
   
   [optional body]
   
   [optional footer]
   ```

3. **Common types**:
   - `feat`: A new feature
   - `fix`: A bug fix
   - `docs`: Documentation changes
   - `style`: Changes that don't affect code meaning
   - `refactor`: Code changes that neither fix a bug nor add a feature
   - `perf`: Performance improvements
   - `test`: Adding or correcting tests
   - `chore`: Changes to build process or auxiliary tools

4. **Reference issues in commits**
   - Use keywords like "Fixes #42" or "Resolves #42"
   - This automatically links the commit to the issue

### Pull Request Workflow

1. **Create descriptive PRs**
   - Clear title that summarizes the change
   - Detailed description explaining the what and why
   - Reference related issues

2. **Keep PRs focused and reasonably sized**
   - Easier to review and less likely to introduce bugs
   - Split large changes into multiple PRs when possible

3. **Request reviews from appropriate team members**

4. **Address review comments promptly**

5. **Merge strategies**
   - Prefer "Squash and merge" for feature branches with multiple small commits
   - Use "Rebase and merge" to preserve commit history when appropriate

### General Tips

1. **Always append `&& sleep 5` to terminal commands** to ensure output is visible

2. **Pull before pushing** to avoid unnecessary merge conflicts

3. **Regularly fetch from remote** to stay updated with team changes

4. **Use `git status` frequently** to check the state of your working directory

5. **Leverage GitHub CLI for efficiency**
   - Create issues and PRs directly from the terminal
   - Check PR status without leaving the command line
   - Quickly clone repositories and check out PRs

## Repository-Specific Notes

- This repository is for an AppScript project for Financial Planning with Dropdowns
- The main remote is: https://github.com/niklas-joh/financial-planning-app-script.git
- The project structure follows AppScript conventions with src/ directory containing the main code

Remember to delete any analysis files in the cline/ directory once analysis is complete, as this directory is only for your reference and is excluded from git.
