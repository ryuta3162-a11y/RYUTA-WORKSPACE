import { redirect } from 'next/navigation';

/** 旧 URL。トップに統一 */
export default function WorkspaceRedirect() {
  redirect('/');
}
