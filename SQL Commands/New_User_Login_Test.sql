SELECT * FROM sys.server_principals WHERE name = 'admin';
ALTER LOGIN [admin] ENABLE;
ALTER LOGIN [admin] WITH DEFAULT_DATABASE = [master];
