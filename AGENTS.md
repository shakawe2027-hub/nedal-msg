<!-- LOVABLE:BEGIN -->
> [!IMPORTANT]
> This project is connected to [Lovable](https://lovable.dev). Avoid rewriting
> published git history — force pushing, or rebasing/amending/squashing commits
> that are already pushed — as it rewrites history on Lovable's side and the
> user will likely lose their project history.
>
> Commits you push to the connected branch sync back to Lovable and show up in
> the editor, so keep the branch in a working state.
<!-- LOVABLE:END -->

- Correspondence data is read and written through Drive-backed server functions using SheetJS; the workbook is the single source of truth without a database.
- Incoming/outgoing pages reuse the shared register and workspace-summary components so presentation stays consistent with reports.
- Health referrals fill the original Drive DOCX XML cells and render it client-side; archive copies in Drive and keep metadata in the workbook to preserve the original form and single source of truth.
