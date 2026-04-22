import 'dotenv/config'; // loads .env into process.env (local dev only)
import express, { Request, Response, NextFunction } from 'express';
import cors from 'cors';
import sql from 'mssql';

// ---------------------------------------------------------------------------
// App initialisation
// ---------------------------------------------------------------------------
const app = express();
const PORT = process.env.PORT || 8080; // Azure App Service injects PORT

// ---------------------------------------------------------------------------
// CORS
// ALLOWED_ORIGINS accepts a comma-separated list of origins, e.g.:
//   ALLOWED_ORIGINS=https://org123.crm.dynamics.com,http://localhost:3002
// ---------------------------------------------------------------------------
const rawOrigins = process.env.ALLOWED_ORIGINS || process.env.ALLOWED_ORIGIN || '';
const allowedOrigins = rawOrigins.split(',').map((o) => o.trim()).filter(Boolean);

const corsOptions: cors.CorsOptions = {
  origin: (origin, callback) => {
    // Allow requests with no origin (e.g. curl, Postman, server-to-server)
    if (!origin) return callback(null, true);
    if (allowedOrigins.includes(origin)) return callback(null, true);
    callback(new Error(`CORS: origin '${origin}' not allowed`));
  },
  methods: ['GET', 'POST', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization'],
};
app.use(cors(corsOptions));
app.use(express.json());

// ---------------------------------------------------------------------------
// Azure SQL connection pool
// Credentials are read exclusively from environment variables — never hardcoded.
// ---------------------------------------------------------------------------
const dbConfig: sql.config = {
  server: process.env.DB_SERVER as string,   // e.g. myserver.database.windows.net
  database: process.env.DB_NAME as string,
  user: process.env.DB_USER as string,
  password: process.env.DB_PASSWORD as string,
  options: {
    encrypt: true,                  // Required for Azure SQL
    trustServerCertificate: false,  // Always false in production
  },
  pool: {
    max: 10,
    min: 0,
    idleTimeoutMillis: 30000,
  },
};

// Create a single reusable connection pool for the lifetime of the process.
const poolPromise: Promise<sql.ConnectionPool> = new sql.ConnectionPool(dbConfig)
  .connect()
  .then((pool) => {
    console.log('Connected to Azure SQL Database');
    return pool;
  })
  .catch((err: Error) => {
    console.error('Database connection failed:', err.message);
    process.exit(1); // Exit so Azure App Service restarts the container
  });

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------
interface TestDataRecord {
  ID: number;
  Title: string;
  Description: string;
}

interface CreateTestDataBody {
  Title: string;
  Description: string;
}

// ---------------------------------------------------------------------------
// Routes
// ---------------------------------------------------------------------------

/**
 * GET /api/testdata
 * Returns all rows from the TestData table.
 */
app.get('/api/testdata', async (_req: Request, res: Response) => {
  try {
    const pool = await poolPromise;
    const result = await pool
      .request()
      .query<TestDataRecord>('SELECT ID, Title, Description FROM dbo.TestData');

    res.status(200).json(result.recordset);
  } catch (err: unknown) {
    // Log the full error server-side; return a clean message to the client.
    console.error('[GET /api/testdata] SQL error:', err);
    res.status(500).json({ message: 'An internal server error occurred.' });
  }
});

/**
 * POST /api/testdata
 * Inserts a new record. Uses parameterized inputs to prevent SQL injection.
 *
 * Expected body: { "Title": "...", "Description": "..." }
 */
app.post('/api/testdata', async (req: Request<{}, {}, CreateTestDataBody>, res: Response) => {
  const { Title, Description } = req.body;

  if (!Title || !Description) {
    res.status(400).json({ message: 'Title and Description are required.' });
    return;
  }

  try {
    const pool = await poolPromise;
    const result = await pool
      .request()
      .input('Title', sql.VarChar, Title)
      .input('Description', sql.VarChar, Description)
      .query<{ NewID: number }>(
        'INSERT INTO dbo.TestData (Title, Description) OUTPUT INSERTED.ID AS NewID VALUES (@Title, @Description)'
      );

    const newId = result.recordset[0].NewID;
    res.status(201).json({ message: 'Record created successfully.', id: newId });
  } catch (err: unknown) {
    console.error('[POST /api/testdata] SQL error:', err);
    res.status(500).json({ message: 'An internal server error occurred.' });
  }
});

/**
 * DELETE /api/testdata/:id
 * Deletes a record by its primary key.
 */
app.delete('/api/testdata/:id', async (req: Request<{ id: string }>, res: Response) => {
  const id = parseInt(req.params.id, 10);
  if (isNaN(id)) {
    res.status(400).json({ message: 'Invalid ID.' });
    return;
  }
  try {
    const pool = await poolPromise;
    await pool
      .request()
      .input('ID', sql.Int, id)
      .query('DELETE FROM dbo.TestData WHERE ID = @ID');
    res.status(200).json({ message: 'Record deleted.' });
  } catch (err: unknown) {
    console.error('[DELETE /api/testdata/:id] SQL error:', err);
    res.status(500).json({ message: 'An internal server error occurred.' });
  }
});

// ---------------------------------------------------------------------------
// Health-check endpoint (used by Azure App Service probes)
// ---------------------------------------------------------------------------
app.get('/health', (_req: Request, res: Response) => {
  res.status(200).json({ status: 'ok' });
});

// ---------------------------------------------------------------------------
// Global error handler (catches anything not caught in route handlers)
// ---------------------------------------------------------------------------
// eslint-disable-next-line @typescript-eslint/no-unused-vars
app.use((err: Error, _req: Request, res: Response, _next: NextFunction) => {
  console.error('[Unhandled error]', err);
  res.status(500).json({ message: 'An unexpected error occurred.' });
});

// ---------------------------------------------------------------------------
// Start
// ---------------------------------------------------------------------------
app.listen(PORT, () => {
  console.log(`Server listening on port ${PORT}`);
});
