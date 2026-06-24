import { HashRouter, Routes, Route } from 'react-router-dom';
import Home from './screens/Home';
import Chat from './screens/Chat';
import Library from './screens/Library';
import Settings from './screens/Settings';

// HashRouter keeps deep links working inside the Capacitor WebView and on
// static PWA hosting without server-side routing.
export default function App() {
  return (
    <HashRouter>
      <Routes>
        <Route path="/" element={<Home />} />
        <Route path="/chat/:chatId" element={<Chat />} />
        <Route path="/library" element={<Library />} />
        <Route path="/settings" element={<Settings />} />
      </Routes>
    </HashRouter>
  );
}
