import { HashRouter, Routes, Route } from 'react-router-dom';
import Home from './screens/Home';
import Chat from './screens/Chat';
import Characters from './screens/Characters';
import CharacterEditor from './screens/CharacterEditor';
import Library from './screens/Library';
import Settings from './screens/Settings';
import ProviderSettings from './screens/ProviderSettings';
import PresetSettings from './screens/PresetSettings';
import Logs from './screens/Logs';

// HashRouter keeps deep links working inside the Capacitor WebView and on
// static PWA hosting without server-side routing.
export default function App() {
  return (
    <HashRouter>
      <Routes>
        <Route path="/" element={<Home />} />
        <Route path="/chat/:chatId" element={<Chat />} />
        <Route path="/characters" element={<Characters />} />
        <Route path="/characters/:characterId" element={<CharacterEditor />} />
        <Route path="/library" element={<Library />} />
        <Route path="/settings" element={<Settings />} />
        <Route path="/settings/provider" element={<ProviderSettings />} />
        <Route path="/settings/preset" element={<PresetSettings />} />
        <Route path="/settings/logs" element={<Logs />} />
      </Routes>
    </HashRouter>
  );
}
