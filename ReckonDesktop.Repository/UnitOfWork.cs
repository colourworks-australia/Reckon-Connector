using System;
using ReckonDesktop.Model;

namespace ReckonDesktop.Repository
{
    public class UnitOfWork : IDisposable
    {
        private readonly ReckonDesktopContext _context = new ReckonDesktopContext();

        // Repositories
        //private GenericRepository<Destination> _destinationRepository;
        //private GenericRepository<FolderMapping> _folderMappingRepository;
        private GenericRepository<Settings> _settingsRepository;
        private GenericRepository<Security> _securityRepository;

        // Repository Implementations
        //public GenericRepository<Destination> DestinationRepository => _destinationRepository ?? (_destinationRepository = new GenericRepository<Destination>(_context));
        //public GenericRepository<FolderMapping> FolderMappingRepository => _folderMappingRepository ?? (_folderMappingRepository = new GenericRepository<FolderMapping>(_context));
        public GenericRepository<Settings> SettingsRepository => _settingsRepository ?? (_settingsRepository = new GenericRepository<Settings>(_context));
        public GenericRepository<Security> SecurityRepository => _securityRepository ?? (_securityRepository = new GenericRepository<Security>(_context));

        public void Save()
        {
            try
            {
                _context.SaveChanges();
            }
            catch (Exception ex)
            {
                Console.WriteLine("There were issues saving to the database: " + ex.Message);
            }
        }

        private bool _disposed = false;

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    _context.Dispose();
                }
            }
            _disposed = true;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
    }
}
