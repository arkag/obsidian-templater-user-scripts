module.exports = async function (tp, jiraQuery, accountAliases, statusToHeadingMap = {}) {
  // Function to update the Kanban file with Jira tasks
  async function updateKanbanWithJiraTasks() {
    // Get the currently opened file
    const activeFile = app.workspace.getActiveFile();
    if (!activeFile) {
      console.error('No file is currently open.');
      return;
    }

    // First, clean up any duplicate cards
    await cleanupDuplicateCards(activeFile);

    // Read the content of the active file (after cleanup)
    let kanbanContent = await app.vault.read(activeFile);

    // Check if the file has the frontmatter variable `autoUpdateKanban: true`
    const frontmatter = app.metadataCache.getFileCache(activeFile)?.frontmatter;
    const isAutoUpdateEnabled = frontmatter?.autoUpdateKanban === true;

    // If the frontmatter variable is not set, add it
    if (!isAutoUpdateEnabled) {
      console.log('First run detected. Adding frontmatter variable.');

      // Check if frontmatter already exists
      const hasExistingFrontmatter = kanbanContent.startsWith('---\n');

      if (hasExistingFrontmatter) {
        // Append `autoUpdateKanban: true` to existing frontmatter
        const frontmatterEndIndex = kanbanContent.indexOf('---\n', 4); // Find the end of the frontmatter block
        if (frontmatterEndIndex !== -1) {
          const existingFrontmatter = kanbanContent.slice(0, frontmatterEndIndex);
          const restOfContent = kanbanContent.slice(frontmatterEndIndex);
          kanbanContent = `${existingFrontmatter}autoUpdateKanban: true\n${restOfContent}`;
        }
      } else {
        kanbanContent = `---\nkanban-plugin: board\nautoUpdateKanban: true\n---\n${kanbanContent}`;
      }

      // Write the updated content back to the active file
      await app.vault.modify(activeFile, kanbanContent);
      return;
    }

    // Fetch Jira tasks based on the query and account aliases
    let jiraTasks = [];
    for (const accountAlias of accountAliases) {
      const tasks = await fetchJiraTasks(jiraQuery, accountAlias);
      jiraTasks = jiraTasks.concat(tasks);
    }

    console.log(`Found ${jiraTasks.length} Jira tasks to process`);

    // Create a map of Jira keys to their current status for quick lookup
    const jiraStatusMap = {};
    jiraTasks.forEach(task => {
      jiraStatusMap[task.key] = task.fields.status.name;
    });

    // First, add new tasks that don't exist in the Kanban
    console.log('Starting to add new tasks to Kanban board');
    const addResult = await addNewTasks(kanbanContent, jiraTasks, statusToHeadingMap, activeFile);
    kanbanContent = addResult.content;

    // Output the number of tasks added
    if (addResult.addedCount > 0) {
      console.log(`Successfully added ${addResult.addedCount} new tasks to the Kanban board`);
    } else {
      console.log('No new tasks were added to the Kanban board');
    }

    // Now, update existing tasks based on their current Jira status
    console.log('Starting to update existing tasks in Kanban board');
    await updateExistingTasks(kanbanContent, jiraStatusMap, statusToHeadingMap, activeFile);

    // After all updates are done, mark items in Complete sections as checked
    kanbanContent = await app.vault.read(activeFile);
    const markedContent = markCompleteItemsAsChecked(kanbanContent);
    if (markedContent !== kanbanContent) {
      console.log('Updated items in Complete sections to be checked');
      await app.vault.modify(activeFile, markedContent);
    }
  }

  // New function to update existing tasks in the Kanban board
  async function updateExistingTasks(kanbanContent, jiraStatusMap, statusToHeadingMap, activeFile) {
    // Extract all task references from the Kanban content
    const taskRegex = /^(- \[[ x]\] \[\[([A-Z]+-\d+)\]\].*?)$/gm;
    let match;
    const tasksToUpdate = [];

    while ((match = taskRegex.exec(kanbanContent)) !== null) {
      const taskKey = match[2];
      if (jiraStatusMap[taskKey]) {
        tasksToUpdate.push({
          key: taskKey,
          status: jiraStatusMap[taskKey],
          fullMatch: match[1],  // The entire task line
          position: match.index,
          length: match[0].length
        });
      }
    }

    console.log(`Found ${tasksToUpdate.length} existing tasks to update`);

    // Extract all headers and their positions
    const headers = [];
    const headerRegex = /^##\s+(.+)$/gm;
    while ((match = headerRegex.exec(kanbanContent)) !== null) {
      headers.push({
        name: match[1],
        position: match.index,
        length: match[0].length
      });
    }

    // Sort headers by position (descending) to avoid position shifts when modifying content
    headers.sort((a, b) => b.position - a.position);

    // Create sections map based on headers
    const sections = {};
    for (let i = 0; i < headers.length; i++) {
      const currentHeader = headers[i];
      const nextHeader = headers[i + 1];
      const sectionStart = currentHeader.position + currentHeader.length;
      const sectionEnd = nextHeader ? nextHeader.position : kanbanContent.length;

      sections[currentHeader.name] = {
        start: sectionStart,
        end: sectionEnd,
        content: kanbanContent.substring(sectionStart, sectionEnd)
      };
    }

    // Process tasks to update (in reverse order to maintain positions)
    tasksToUpdate.sort((a, b) => b.position - a.position);

    // Create a copy of the content to work with
    let updatedContent = kanbanContent;
    let tasksUpdated = 0;

    // First, collect all tasks to move by target section
    const tasksByTargetSection = {};

    for (const task of tasksToUpdate) {
      // Find which section the task is currently in
      let currentSection = null;
      for (const [headerName, section] of Object.entries(sections)) {
        if (task.position >= section.start && task.position < section.end) {
          currentSection = headerName;
          break;
        }
      }

      if (!currentSection) {
        console.log(`Could not determine current section for task ${task.key}`);
        continue;
      }

      // Determine target section based on Jira status
      const targetHeader = statusToHeadingMap[task.status] || task.status;

      // If task is already in the correct section, skip it
      if (currentSection === targetHeader) {
        console.log(`Task ${task.key} is already in the correct section: ${targetHeader}`);
        continue;
      }

      // Add task to the collection for its target section
      if (!tasksByTargetSection[targetHeader]) {
        tasksByTargetSection[targetHeader] = [];
      }

      tasksByTargetSection[targetHeader].push({
        key: task.key,
        line: `- [ ] [[${task.key}]]`,
        currentPosition: task.position,
        currentLength: task.length,
        currentSection: currentSection
      });

      console.log(`Will move task ${task.key} from "${currentSection}" to "${targetHeader}"`);
    }

    // Track which sections have had tasks removed for later cleanup
    const sectionsWithRemovedTasks = new Set();

    // Now remove all tasks from their original positions (in reverse order)
    const tasksToRemove = Object.values(tasksByTargetSection)
      .flat()
      .sort((a, b) => b.currentPosition - a.currentPosition);

    for (const task of tasksToRemove) {
      // Remove task from current position
      updatedContent = updatedContent.slice(0, task.currentPosition) +
        updatedContent.slice(task.currentPosition + task.currentLength);
      sectionsWithRemovedTasks.add(task.currentSection);
      tasksUpdated++;
    }

    // Now add all tasks to their target sections
    for (const [targetHeader, tasks] of Object.entries(tasksByTargetSection)) {
      if (tasks.length === 0) continue;

      // Find the target header position in the updated content
      const targetHeaderPos = updatedContent.indexOf(`## ${targetHeader}`);

      if (targetHeaderPos === -1) {
        // If target header doesn't exist, create it and add all tasks
        const tasksText = tasks.map(t => t.line).join('\n');
        updatedContent += `\n\n## ${targetHeader}\n${tasksText}`;
      } else {
        // Find where to insert in the target section
        const headerEndPos = updatedContent.indexOf('\n', targetHeaderPos);

        if (headerEndPos === -1) {
          // If no newline after header, append to the end
          const tasksText = tasks.map(t => t.line).join('\n');
          updatedContent += `\n${tasksText}`;
        } else {
          // Check if there's a "Complete" subheader in this section
          const sectionEndPos = updatedContent.indexOf('\n## ', headerEndPos);
          const sectionContent = sectionEndPos !== -1 ?
            updatedContent.substring(headerEndPos, sectionEndPos) :
            updatedContent.substring(headerEndPos);

          const completePos = sectionContent.indexOf('\n**Complete**');

          if (completePos !== -1) {
            // Insert after the "Complete" subheader
            const completeLinePos = headerEndPos + completePos;
            const completeLineEndPos = updatedContent.indexOf('\n', completeLinePos + 1);

            if (completeLineEndPos !== -1) {
              // Insert tasks after the "Complete" line
              const tasksText = tasks.map(t => t.line).join('\n');
              updatedContent = updatedContent.slice(0, completeLineEndPos + 1) +
                tasksText + '\n' +
                updatedContent.slice(completeLineEndPos + 1);
            } else {
              // "Complete" is the last line in the file
              const tasksText = tasks.map(t => t.line).join('\n');
              updatedContent += `\n${tasksText}`;
            }
          } else {
            // No "Complete" subheader, insert tasks after the header line
            const tasksText = tasks.map(t => t.line).join('\n');
            updatedContent = updatedContent.slice(0, headerEndPos + 1) +
              tasksText + '\n' +
              updatedContent.slice(headerEndPos + 1);
          }
        }
      }
    }

    // Clean up excessive whitespace in the document
    if (tasksUpdated > 0) {
      updatedContent = cleanupWhitespace(updatedContent);
    }

    if (tasksUpdated > 0) {
      console.log(`Updated ${tasksUpdated} tasks in the Kanban board`);
      await app.vault.modify(activeFile, updatedContent);
    } else {
      console.log('No tasks needed to be moved');
    }
  }

  // Function to fetch Jira tasks
  async function fetchJiraTasks(query, accountAlias) {
    // Use the Obsidian Jira Issue plugin to search for tasks
    const jiraIssues = await searchJiraIssues(query, accountAlias);
    return jiraIssues;
  }

  // Function to search Jira issues
  async function searchJiraIssues(query, accountAlias) {
    // Ensure the Jira Issue plugin is loaded
    if (typeof $ji === 'undefined') {
      console.error('Obsidian Jira Issue plugin is not loaded.');
      return [];
    }

    const account = $ji.account.getAccountByAlias(accountAlias);

    // Call the plugin's getSearchResults method (correct method according to documentation)
    const searchResults = await $ji.defaulted.getSearchResults(query, { account });
    // Format the results into a consistent structure
    return searchResults.issues.map(issue => ({
      key: issue.key,
      fields: {
        summary: issue.fields.summary,
        status: { name: issue.fields.status.name }
      }
    }));
  }

  // Function to extract headers from the file content
  function extractHeadersFromContent(content) {
    const headerRegex = /##\s+(.+)\n/g;
    const headers = [];
    let match;
    while ((match = headerRegex.exec(content)) !== null) {
      headers.push(match[1]);
    }
    return headers;
  }

  // Function to check if a file exists
  async function fileExists(fileName) {
    const file = app.vault.getAbstractFileByPath(fileName);
    return file !== null;
  }

  // Function to add new Jira tasks to the Kanban board
  async function addNewTasks(kanbanContent, jiraTasks, statusToHeadingMap, activeFile) {
    console.log('Starting to add new tasks to Kanban board');

    // Extract existing task keys from the Kanban content
    const existingTaskRegex = /\[\[([A-Z]+-\d+)\]\]/g;
    const existingTaskKeys = new Set();
    let match;

    while ((match = existingTaskRegex.exec(kanbanContent)) !== null) {
      existingTaskKeys.add(match[1]);
    }

    // Identify new tasks that don't exist in the Kanban yet
    const newTasks = jiraTasks.filter(task => !existingTaskKeys.has(task.key));
    console.log(`Found ${newTasks.length} new tasks to add to Kanban`);

    if (newTasks.length === 0) {
      return { content: kanbanContent, addedCount: 0 };
    }

    // Group tasks by their target section
    const tasksBySection = {};

    for (const task of newTasks) {
      const status = task.fields.status.name;
      const targetHeader = statusToHeadingMap[status] || status;

      if (!tasksBySection[targetHeader]) {
        tasksBySection[targetHeader] = [];
      }

      tasksBySection[targetHeader].push({
        key: task.key,
        line: `- [ ] [[${task.key}]]`
      });
    }

    let updatedContent = kanbanContent;
    let tasksAdded = 0;

    // Add tasks to each section
    for (const [targetHeader, tasks] of Object.entries(tasksBySection)) {
      if (tasks.length === 0) continue;

      // Find the target header position
      const targetHeaderPos = updatedContent.indexOf(`## ${targetHeader}`);

      if (targetHeaderPos === -1) {
        // If target header doesn't exist, create it and add all tasks
        const tasksText = tasks.map(t => t.line).join('\n');
        updatedContent += `\n\n## ${targetHeader}\n${tasksText}`;
      } else {
        // Find where to insert in the target section
        const headerEndPos = updatedContent.indexOf('\n', targetHeaderPos);

        if (headerEndPos === -1) {
          // If no newline after header, append to the end
          const tasksText = tasks.map(t => t.line).join('\n');
          updatedContent += `\n${tasksText}`;
        } else {
          // Check if there's a "Complete" line after the header
          const nextLineStart = headerEndPos + 1;
          const nextLineEnd = updatedContent.indexOf('\n', nextLineStart);
          const nextLine = nextLineEnd !== -1 ?
            updatedContent.substring(nextLineStart, nextLineEnd) :
            updatedContent.substring(nextLineStart);

          if (nextLine.trim() === '**Complete**') {
            // Insert after the "Complete" line
            const tasksText = tasks.map(t => t.line).join('\n');
            updatedContent = updatedContent.slice(0, nextLineEnd + 1) +
              tasksText + '\n' +
              updatedContent.slice(nextLineEnd + 1);
          } else {
            // Insert tasks after the header line
            const tasksText = tasks.map(t => t.line).join('\n');
            updatedContent = updatedContent.slice(0, headerEndPos + 1) +
              tasksText + '\n' +
              updatedContent.slice(headerEndPos + 1);
          }
        }
      }

      tasksAdded += tasks.length;
    }

    // Clean up excessive whitespace in the document
    if (tasksAdded > 0) {
      updatedContent = cleanupWhitespace(updatedContent);
    }

    console.log(`Added ${tasksAdded} new tasks to the Kanban board`);
    await app.vault.modify(activeFile, updatedContent);
    return { content: updatedContent, addedCount: tasksAdded };
  }

  // Function to clean up duplicate cards in the Kanban board
  async function cleanupDuplicateCards(activeFile) {
    console.log('Starting to clean up duplicate cards in Kanban board');

    // Read the content of the active file
    let kanbanContent = await app.vault.read(activeFile);

    // Extract all task references from the Kanban content
    const taskRegex = /^(- \[[ x]\] \[\[([A-Z]+-\d+)\]\].*?)$/gm;
    let match;
    const taskOccurrences = {};
    const taskPositions = [];

    while ((match = taskRegex.exec(kanbanContent)) !== null) {
      const taskKey = match[2];
      const position = match.index;
      const length = match[0].length;

      if (!taskOccurrences[taskKey]) {
        taskOccurrences[taskKey] = 1;
      } else {
        taskOccurrences[taskKey]++;
        // Store position and length of duplicate tasks (not the first occurrence)
        taskPositions.push({ key: taskKey, position, length });
      }
    }

    // Count duplicates
    const duplicateKeys = Object.keys(taskOccurrences).filter(key => taskOccurrences[key] > 1);
    console.log(`Found ${duplicateKeys.length} tasks with duplicates`);

    if (duplicateKeys.length === 0) {
      console.log('No duplicate cards to clean up');
      return;
    }

    // Sort positions in reverse order to avoid position shifts when removing content
    taskPositions.sort((a, b) => b.position - a.position);

    let updatedContent = kanbanContent;
    let removedCount = 0;

    // Remove duplicate tasks
    for (const task of taskPositions) {
      console.log(`Removing duplicate of task ${task.key} at position ${task.position}`);

      // Remove task from current position
      updatedContent = updatedContent.slice(0, task.position) +
        updatedContent.slice(task.position + task.length);

      removedCount++;
    }

    if (removedCount > 0) {
      // Clean up any excessive whitespace created by removing duplicates
      updatedContent = cleanupWhitespace(updatedContent);

      // Mark items in Complete sections as checked
      updatedContent = markCompleteItemsAsChecked(updatedContent);

      console.log(`Removed ${removedCount} duplicate cards from the Kanban board`);
      await app.vault.modify(activeFile, updatedContent);
    } else {
      // Even if no duplicates were removed, still mark Complete items as checked
      updatedContent = markCompleteItemsAsChecked(kanbanContent);
      if (updatedContent !== kanbanContent) {
        console.log('Updated items in Complete sections to be checked');
        await app.vault.modify(activeFile, updatedContent);
      }
    }
  }

  // Helper function to clean up whitespace in the document
  function cleanupWhitespace(content) {
    let cleanedContent = content;

    // Normalize all newlines to \n
    cleanedContent = cleanedContent.replace(/\r\n/g, '\n');

    // Ensure exactly one newline between list items
    cleanedContent = cleanedContent.replace(/(- \[[ x]\] .+?)(\n+)(- \[[ x]\])/g, '$1\n$3');

    // Ensure exactly one newline between "Complete" marker and the first list item
    cleanedContent = cleanedContent.replace(/(\*\*Complete\*\*)(\n+)(- \[)/g, '$1\n$3');

    // Ensure exactly one newline between header and "Complete" marker
    cleanedContent = cleanedContent.replace(/(##\s+.+?)(\n+)(\*\*Complete\*\*)/g, '$1\n$3');

    // Ensure exactly two newlines between header and first list item (when no Complete marker)
    cleanedContent = cleanedContent.replace(/(##\s+.+?)(\n+)(- \[)/g, '$1\n\n$3');

    // Ensure exactly two newlines between the end of a list and the next section header
    cleanedContent = cleanedContent.replace(/(- \[[ x]\] .+?)(\n+)(##\s+)/g, '$1\n\n$3');

    // Clean up whitespace at the beginning and end of the file
    cleanedContent = cleanedContent.trim() + '\n';

    return cleanedContent;
  }

  // Add this function after cleanupWhitespace
  function markCompleteItemsAsChecked(content) {
    // Find all sections with a **Complete** subheader
    const completeHeaderRegex = /##\s+(.+?)\n\*\*Complete\*\*\n([\s\S]+?)(?=\n##|$)/g;
    let match;
    let updatedContent = content;

    while ((match = completeHeaderRegex.exec(content)) !== null) {
      const sectionName = match[1].trim();
      const sectionContent = match[2];

      // Replace all unchecked items with checked items in this section
      const checkedSectionContent = sectionContent.replace(/- \[ \] \[\[/g, '- [x] [[');

      // Replace the section in the content
      updatedContent = updatedContent.replace(match[0], `## ${sectionName}\n**Complete**\n${checkedSectionContent}`);
    }

    return updatedContent;
  }

  // Expose functions to Templater
  return {
    updateKanbanWithJiraTasks,
    fetchJiraTasks,
    searchJiraIssues,
    extractHeadersFromContent,
    fileExists,
    addNewTasks,
    cleanupDuplicateCards
  };
};
