import { useState, useEffect } from 'react';
import { Track } from '../domain/entities/Track';
import { Playlist } from '../domain/entities/Playlist';
import { useMusicRepository } from '../presentation/hooks/useMusicRepository';
import { useAudioPlayerContext } from '../presentation/context/AudioPlayerContext';
import { SearchBar } from '../presentation/components/Search/SearchBar';
import { TrackList } from '../presentation/components/TrackList/TrackList';
import { PlaylistGrid } from '../presentation/components/Playlists/PlaylistGrid';
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs';
import { Loader2, Music } from 'lucide-react';
import { toast } from '@/hooks/use-toast';

const Index = () => {
  const repository = useMusicRepository();
  const { play, currentTrack, playerState, pause, resume } = useAudioPlayerContext();
  
  const [tracks, setTracks] = useState<Track[]>([]);
  const [userPlaylists, setUserPlaylists] = useState<Playlist[]>([]);
  const [selectedPlaylist, setSelectedPlaylist] = useState<Playlist | null>(null);
  
  const [isLoadingTracks, setIsLoadingTracks] = useState(false);
  const [isLoadingPlaylists, setIsLoadingPlaylists] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');

  // Load initial tracks and playlists
  useEffect(() => {
    loadInitialData();
  }, []);

  const loadInitialData = async () => {
    try {
      setIsLoadingTracks(true);
      setIsLoadingPlaylists(true);

      const [tracksData, userData] = await Promise.all([
        repository.searchTracks(''),
        repository.getUserPlaylists(),
      ]);

      setTracks(tracksData);
      setUserPlaylists(userData);
    } catch (error) {
      console.error('Failed to load data:', error);
      toast({
        title: 'Error',
        description: 'Failed to load music data',
        variant: 'destructive',
      });
    } finally {
      setIsLoadingTracks(false);
      setIsLoadingPlaylists(false);
    }
  };

  const handleSearch = async (query: string) => {
    try {
      setIsLoadingTracks(true);
      setSearchQuery(query);
      setSelectedPlaylist(null);
      const results = await repository.searchTracks(query);
      setTracks(results);
    } catch (error) {
      console.error('Search failed:', error);
      toast({
        title: 'Error',
        description: 'Search failed',
        variant: 'destructive',
      });
    } finally {
      setIsLoadingTracks(false);
    }
  };

  const handleTrackSelect = (track: Track) => {
    if (currentTrack?.id === track.id) {
      // Toggle play/pause for the same track
      if (playerState === 'playing') {
        pause();
      } else {
        resume();
      }
    } else {
      // Play new track
      play(track);
    }
  };

  const handlePlaylistSelect = (playlist: Playlist) => {
    setSelectedPlaylist(playlist);
    setTracks(playlist.tracks);
    setSearchQuery('');
  };

  const displayedTracks = selectedPlaylist ? selectedPlaylist.tracks : tracks;

  return (
    <div className="min-h-screen bg-background pb-32">
      {/* Header */}
      <header className="border-b border-border bg-card/50 backdrop-blur-sm sticky top-0 z-10">
        <div className="max-w-7xl mx-auto px-4 py-4 md:py-6">
          <div className="flex flex-col md:flex-row md:items-center gap-4 md:gap-6">
            <div className="flex items-center gap-3">
              <Music className="h-8 w-8 text-primary" />
              <h1 className="text-2xl md:text-3xl font-bold">Music Streaming</h1>
            </div>
            <div className="flex-1 md:max-w-md">
              <SearchBar onSearch={handleSearch} isLoading={isLoadingTracks} />
            </div>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="max-w-7xl mx-auto px-4 py-8">
        <Tabs defaultValue="search" className="w-full">
          <div className="flex flex-col md:flex-row gap-6">
            {/* Sidebar Navigation */}
            <aside className="md:w-48 flex-shrink-0">
              <TabsList className="flex flex-row md:flex-col h-auto w-full gap-1">
                <TabsTrigger value="search" className="flex-1 md:w-full md:justify-start">Главная</TabsTrigger>
                <TabsTrigger value="user" className="flex-1 md:w-full md:justify-start">Ваши плейлисты</TabsTrigger>
              </TabsList>
            </aside>

            {/* Content Area */}
            <div className="flex-1 min-w-0">
              <TabsContent value="search" className="mt-0 space-y-6">
                {selectedPlaylist && (
                  <div className="mb-6">
                    <h2 className="text-2xl font-bold mb-2">{selectedPlaylist.name}</h2>
                    {selectedPlaylist.description && (
                      <p className="text-muted-foreground">{selectedPlaylist.description}</p>
                    )}
                  </div>
                )}

                {isLoadingTracks ? (
                  <div className="flex items-center justify-center py-12">
                    <Loader2 className="h-8 w-8 animate-spin text-primary" />
                  </div>
                ) : (
                  <TrackList
                    tracks={displayedTracks}
                    onTrackSelect={handleTrackSelect}
                    currentTrackId={currentTrack?.id}
                  />
                )}
              </TabsContent>

              <TabsContent value="user" className="mt-0">
                {isLoadingPlaylists ? (
                  <div className="flex items-center justify-center py-12">
                    <Loader2 className="h-8 w-8 animate-spin text-primary" />
                  </div>
                ) : (
                  <PlaylistGrid
                    playlists={userPlaylists}
                    onPlaylistSelect={handlePlaylistSelect}
                  />
                )}
              </TabsContent>
            </div>
          </div>
        </Tabs>
      </main>
    </div>
  );
};

export default Index;
