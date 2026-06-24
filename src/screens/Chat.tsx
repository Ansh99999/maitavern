import { useParams } from 'react-router-dom';
import Placeholder from './Placeholder';

// Chat interface (docs/01): 4 display styles, avatar viewer, composer. Phase 1.
export default function Chat() {
  const { chatId } = useParams();
  return <Placeholder title="Chat" note={`Chat interface for "${chatId}" — built in Phase 1.`} />;
}
