import { Link } from 'react-router-dom';
import Icon from '@/components/Icon';

/** Shared themed placeholder for Phase 0 screens not yet built out. */
export default function Placeholder({ title, note }: { title: string; note: string }) {
  return (
    <div className="min-h-full bg-bg text-text font-ui">
      <header className="flex items-center gap-3 px-5 pt-4 pb-3 border-b border-border">
        <Link to="/" aria-label="Back" className="text-muted">
          <Icon name="chevronLeft" size={22} />
        </Link>
        <h1 className="text-lg font-semibold">{title}</h1>
      </header>
      <main className="px-5 py-8 text-muted">{note}</main>
    </div>
  );
}
