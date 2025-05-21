function deleteOldLabeledEmails() {
  const labelName = 'Deleted';
  const label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    Logger.log('Label "' + labelName + '" not found.');
    return;
  }

  const threads = label.getThreads();
  const now = new Date();
  const cutoffDate = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000); // 30 days ago

  let deletedCount = 0;

  for (const thread of threads) {
    const lastMessageDate = thread.getLastMessageDate();
    if (lastMessageDate < cutoffDate) {
      thread.moveToTrash();
      deletedCount++;
    }
  }

  Logger.log('Moved ' + deletedCount + ' threads to Trash.');
}
