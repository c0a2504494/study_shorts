function includeStudyApp() {
  const source = HtmlService.createHtmlOutputFromFile('App').getContent();
  const styleMatch = source.match(/<style>([\s\S]*?)<\/style>/i);
  const bodyMatch = source.match(/<body[^>]*>([\s\S]*?)<\/body>/i);

  if (!bodyMatch) {
    throw new Error('App.html body could not be loaded.');
  }

  const styles = styleMatch ? `<style>${styleMatch[1]}</style>` : '';
  return `${styles}${bodyMatch[1]}`;
}
