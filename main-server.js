import express from 'express';
import { MongoClient, ObjectId } from 'mongodb';
import XLSX from 'xlsx';
import cors from 'cors';

const app = express();
app.use(express.json({ limit: '20mb' }));
app.use(express.urlencoded({ extended: true, limit: '20mb' }));
app.use(express.json());
app.use(cors());

const uri = 'mongodb://localhost:27017';
const client = new MongoClient(uri);

const DB_NAME = "docuflow4"

app.get("/", (req,res) => res.send("Welcome to DocuFlow AI"))

app.post('/docuflow/citi_data', async (req, res) => {
    await client.connect();
    const db = client.db(DB_NAME);

    const jobsCollection = db.collection('citi_data');
    await jobsCollection.insertMany(req.body);

    // Handle document processing
    res.status(201).json({ success: true });
});

app.post('/docuflow/jobs', async (req, res) => {
    await client.connect();

    const db = client.db(DB_NAME);

    const jobsCollection = db.collection('jobs');
    await jobsCollection.insertMany(req.body);

    // Handle document processing
    res.status(201).json({ success: true });
});

app.post('/docuflow/jobs/update', async (req, res) => {
    await client.connect();

    const db = client.db(DB_NAME);

    const collection = db.collection('jobs');
    const data = req.body

    const jobId = data.jobId;
    const updatedData = data.updatedData;

    let query;
    if (isValidObjectId(jobId)) {
      // If it's a valid ObjectId, search by _id
      query = { _id: new ObjectId(jobId) };
    } else {
      // Try with batchId for non-ObjectId values
      query = { batchId: jobId };
    }

    const result = await collection.updateOne(
      query,
      { 
        $set: { 
          data: updatedData,
          updatedAt: new Date()
        } 
      }
    );

    if (result.matchedCount === 0) {
      
      // If no match with the first query and it's not a valid ObjectId,
      // try an additional string comparison on _id
      if (!isValidObjectId(jobId)) {
        const stringResult = await collection.updateOne(
          { _id: jobId },
          { 
            $set: { 
              data: updatedData,
              updatedAt: new Date()
            } 
          }
        );
        
        if (stringResult.matchedCount === 0) {
          return res.json({ error: 'Job not found' }, { status: 404 });
        }
      } else {
        return res.json({ error: 'Job not found' }, { status: 404 });
      }
    }

    return res.json({ 
      message: 'Job updated successfully!',
      modifiedCount: result.modifiedCount
    }, { status: 200 });
});

app.post('/docuflow/jobs/update-job-aggrement', async (req, res) => {
    await client.connect();
    const db = client.db(DB_NAME);
    const collection = db.collection('jobs');

    const data= req.body

    const jobId = data.jobId;
    const agreementNumber = data.agreementNumber;
    const updatedEntry = data.updatedEntry;

    let query;
    if (isValidObjectId(jobId)) {
      query = { _id: new ObjectId(jobId) };
    } else {
      query = { _id: jobId };
    }
    
    // First, get the current job document
    const job = await collection.findOne(query);
    
    if (!job) {
      console.error(`Job not found: ${jobId}`);
      return res.json({ error: 'Job not found' }, { status: 404 });
    }
    
    // Normalize agreement number for comparison
    const agreementNumAsString = String(agreementNumber).trim();
    const agreementNumAsInt = parseInt(agreementNumAsString, 10);
    
    // Initialize data array if it doesn't exist
    if (!job.data) {
      job.data = [];
    }
    
    // Find if there's an existing entry for this agreement number
    const existingIndex = job.data.findIndex((entry) => {
      if (!entry || !entry['Agreement Number']) return false;
      
      const entryAgreement = String(entry['Agreement Number']).trim();
      return entryAgreement === agreementNumAsString || 
             parseInt(entryAgreement, 10) === agreementNumAsInt;
    });
    
    // Updated data array
    const updatedData = [...job.data];
    
    if (existingIndex >= 0) {
      updatedData[existingIndex] = updatedEntry;
    } else {
      updatedData.push(updatedEntry);
    }
    
    // Update the document
    const result = await collection.updateOne(
      query,
      { 
        $set: { 
          data: updatedData,
          updatedAt: new Date()
        } 
      }
    );
    
    if (result.matchedCount === 0) {
      return res.json({ error: 'Failed to update job' }, { status: 500 });
    }
    
    console.log(`Job updated successfully: ${result.modifiedCount} document(s) modified`);
    
    return res.json({ 
      message: 'Job updated successfully!',
      modifiedCount: result.modifiedCount,
      jobId: job._id
    }, { status: 200 });
})

app.post('/docuflow/jobs/upload', async (req, res) => {
  const rows = req.body;
  console.log('Rows received:', rows);

  try {
    await client.connect();
    const db = client.db(DB_NAME);
    const collection = db.collection('jobs');

    // Create individual documents for each row with a common batch ID
    const batchId = new Date().getTime().toString();
    const documents = rows.map((row) => ({
      ...row,
      batchId,
      createdAt: new Date(),
      updatedAt: new Date(),
    }));

    console.log(`Inserting ${documents.length} documents into MongoDB`); // Debugging
    await collection.insertMany(documents);
    return res.status(200).json({ message: 'Data uploaded successfully!' });
  } catch (error) {
    console.error('Failed to upload data:', error); // Debugging
    return res.status(500).json({ error: 'Failed to upload data' });
  } finally {
    await client.close();
  }
});

app.get('/docuflow/jobs', async (req, res) => {
    await client.connect();
    console.log('✅ Connected to MongoDB');

    const db = client.db(DB_NAME);
    const collection = db.collection('jobs');

    const pipeline = [
      {
        $group: {
          _id: "$batchId",
          count: { $sum: 1 },
          createdAt: { $first: "$createdAt" },
          updatedAt: { $first: "$updatedAt" },
          // Sample data for display
          sampleData: { $first: { 
            "Agreement Number": "$Agreement Number", 
            "Customer Name": "$Customer Name" 
          }}
        }
      },
      {
        $project: {
          _id: 1,
          count: 1,
          createdAt: 1,
          updatedAt: 1,
          batchId: "$_id",
          // Empty data array to maintain compatibility
          data: { $literal: [] }
        }
      }
    ];
    
    const jobSummaries = await collection.aggregate(pipeline).toArray();

    console.log(jobSummaries)

    // Handle document processing
    res.json({ success: true, body: jobSummaries });
});

app.get('/docuflow/jobs/stats', async (req, res) => {
    const batchId = req.query.batchId

    console.log({})

    await client.connect();
    console.log('✅ Connected to MongoDB');

    const db = client.db(DB_NAME);
    const collection = db.collection('jobs');

    const jobs = await collection.find({ 
      batchId
    }).toArray();
    
    // Count matched items in data arrays
    let matchedCount = 0;
    
    // First check if there are any matched entries in data arrays
    for (const job of jobs) {
      if (Array.isArray(job.data)) {
        const matchedEntries = job.data.filter((item) => item.status === 'Matched');
        matchedCount += matchedEntries.length;
      }
      
      // Also check if the job itself is marked as matched at the top level
      if (job.status === 'Matched') {
        matchedCount += 1;
      }
    }

    res.json({ success: true, matchedCount });
});

app.get('/docuflow/jobs/check-job-aggrement', async (req, res) => {
    console.log(req.query)
    const agreementNumber = req.query.agreementNumber;
    const jobId = req.query.jobId;

    await client.connect();

    console.log({ agreementNumber, jobId })

    const db = client.db(DB_NAME);
    const collection = db.collection('jobs');

    const agreementNumAsString = String(agreementNumber).trim();
    const agreementNumAsInt = parseInt(agreementNumAsString, 10);

    const query = {
      $or: [
        { 'Agreement Number': agreementNumAsString },
        { 'Agreement Number': agreementNumAsInt }
      ]
    };

    console.log(query)
    
    let anyJob = await collection.findOne(query);

    console.log({anyJob})
    
    // If not found at top level, check in data arrays
    if (!anyJob) {
      console.log('Not found at top level, checking in data arrays...');
      anyJob = await collection.findOne({
        'data': {
          $elemMatch: {
            'Agreement Number': { $in: [agreementNumAsString, agreementNumAsInt] }
          }
        }
      });
    }
    
    // If we found a matching agreement anywhere, return that it exists
    if (anyJob) {
      console.log(`Match found for agreement number: ${agreementNumber}, job ID: ${anyJob._id}`);
      
      let matchedEntry;
      
      // If found in the data array, extract that specific entry
      if (Array.isArray(anyJob.data)) {
        matchedEntry = anyJob.data.find((item) => {
          if (!item || !item['Agreement Number']) return false;
          
          const itemAgreement = String(item['Agreement Number']).trim();
          return itemAgreement === agreementNumAsString || 
                 parseInt(itemAgreement, 10) === agreementNumAsInt;
        });
      }
      
      // If not found in the data array, the agreement is at the top level
      if (!matchedEntry) {
        matchedEntry = {
          'Agreement Number': anyJob['Agreement Number'],
          'Customer Name': anyJob['Customer Name']
        };
      }
      
      // Determine if this is the requested job or a different one
      let isRequestedJob = false;
      if (jobId) {
        if (isValidObjectId(jobId)) {
          // For ObjectId-based job IDs
          const jobObjectId = new ObjectId(jobId);
          isRequestedJob = anyJob._id.toString() === jobObjectId.toString();
        } else {
          // For string-based identifiers (like batch IDs)
          isRequestedJob = anyJob.batchId === jobId || anyJob._id.toString() === jobId;
          console.log(`Using string comparison for job ID: ${jobId}, matches: ${isRequestedJob}`);
        }
      }
      
      return res.json({ 
        exists: true,
        matchedEntry: matchedEntry,
        jobId: anyJob._id,
        isRequestedJob: isRequestedJob,
        fullDocument: anyJob
      });
    }

    return res.json({ 
      exists: false,
      matchedEntry: null
    });
});

app.get('/docuflow/jobs/check-match', async (req, res) => {
    const agreementNumber = req.query.agreementNumber

    await client.connect();
    console.log('✅ Connected to MongoDB');

    const db = client.db(DB_NAME);
    const collection = db.collection('jobs');

    const matchedEntry = await collection.findOne({
      'data.Agreement Number': agreementNumber,
      'data.status': 'Matched'
    });

    return res.json({ 
      isMatched: !!matchedEntry,
      matchedEntry: matchedEntry?.data.find((row) => row['Agreement Number'] === agreementNumber)
    }, { status: 200 });
})

app.get('/docuflow/jobs/count-aggrements', async (req, res) => {
     await client.connect();
    console.log('✅ Connected to MongoDB');

    const db = client.db(DB_NAME);
    const collection = db.collection('jobs');

    const totalAgreements = await collection.aggregate([
      { $unwind: "$data" },
      { $count: "totalAgreements" }
    ]).toArray();

    return res.json({ 
      totalAgreements: totalAgreements[0]?.totalAgreements || 0 
    });
})

app.get('/docuflow/jobs/count-citi', async (req, res) => {
    const jobId = req.query.jobId

     await client.connect();
    console.log('✅ Connected to MongoDB');

    const db = client.db(DB_NAME);
    const collection = db.collection('citi_data');

    const query = jobId ? { jobId } : {};
    const citiCount = await collection.countDocuments(query);

    return res.json({ 
      citiCount 
    });
})

app.get('/docuflow/jobs/count-matched', async (req, res) => {
     await client.connect();
    console.log('✅ Connected to MongoDB');

    const db = client.db(DB_NAME);
    const collection = db.collection('jobs');

    const matchedCount = await collection.aggregate([
      { $unwind: "$data" },
      { $match: { "data.status": "Matched" } },
      { $count: "matchedCount" }
    ]).toArray();

    return res.json({ 
        success: true,
        matchedCount: matchedCount[0]?.matchedCount || 0 
    });
})

app.post('/ocrdatabase/quarantine_data', async (req, res) => {
    await client.connect();
    console.log('✅ Connected to MongoDB');

    docuFlowDB = client.db('docuflow');
    ocrDatabaseDB = client.db('ocr_database1')
  // Handle OCR scan
  res.json({ message: 'ocr_database: scan endpoint hit', body: req.body });
});

app.post('/ocrdatabase/agreements', async (req, res) => {
    await client.connect();
    console.log('✅ Connected to MongoDB');

    docuFlowDB = client.db('docuflow4');
    ocrDatabaseDB = client.db('ocr_database2')
  // Handle storing OCR result
  res.json({ message: 'ocr_database: store endpoint hit', body: req.body });
});

app.get("/docuflow/citi_data", async (req, res) => {
  const jobId = req.query.jobId;

  try {
    await client.connect();
    const db = client.db(DB_NAME);
    const collection = db.collection('citi_data');
    
    // If jobId is provided, filter by it
    const query = jobId ? { jobId } : {};
    console.log(`Fetching Citi data with query:`, query);
    
    const data = await collection.find(query).toArray();
    console.log(`Found ${data.length} Citi entries`);
    
    return res.status(200).json(data);
  } catch (error) {
    console.error('Error fetching Citi data:', error);
    return res.status(500).json({ error: 'Failed to fetch Citi data' });
  } finally {
    if (client) {
      await client.close();
    }
  }
})

app.post("/docuflow/citi_data", async (req, res) => {
  client = new MongoClient(process.env.MONGODB_URI);
    await client.connect();
    const db = client.db(DB_NAME);
    const collection = db.collection('citi_data');

    const data  = req.body
    const existingEntry = await collection.findOne({
      'Agreement Number': data['Agreement Number'],
      jobId: data.jobId
    });

    if (existingEntry) {
      console.log('Citi entry already exists, skipping save');
      return res.json({ success: true, message: 'Entry already exists' }, { status: 200 });
    }

    const citiEntry = {
      'Agreement Number': data['Agreement Number'],
      'Customer Name': data['Customer Name'],
      date: data.date,
      status: data.status || 'Matched',
      storageBoxNumber: data.storageBoxNumber,
      matchedImage: data.matchedImage,
      matchedDate: data.matchedDate,
      scanBoxNumber: data.scanBoxNumber,
      barcode_number: data.barcode_number || '',
      scanned_barcode: data.scanned_barcode || data.barcode_number || '', // Ensure scanned_barcode field is included
      password: data.password,
      jobId: data.jobId,
      createdAt: new Date(),
      remarks: data.remarks
    };
    
    // Log what's being inserted to database
    console.log("Inserting into database:", citiEntry);

    const result = await collection.insertOne(citiEntry);
    return res.json({ success: true, insertedId: result.insertedId });
})

app.post("/docuflow/quarantine_data", async (req, res) => {
    await client.connect();
    const db = client.db(DB_NAME);
    const collection = db.collection('quarantine_data');

    const data  = req.body
     const existingEntry = await collection.findOne({
      'Agreement Number': data['Agreement Number'],
      jobId: data.jobId
    });

    if (existingEntry) {
      console.log('Quarantine entry already exists, skipping save');
      return res.json({ success: true, message: 'Entry already exists' }, { status: 200 });
    }

    // Explicitly create object with all fields including barcode
    const quarantineEntry = {
      'Agreement Number': data['Agreement Number'],
      'Customer Name': data['Customer Name'],
      date: data.date,
      status: data.status || 'Matched',
      storageBoxNumber: data.storageBoxNumber,
      matchedImage: data.matchedImage,
      matchedDate: data.matchedDate,
      scanBoxNumber: data.scanBoxNumber,
      barcode_number: data.barcode_number || '',
      scanned_barcode: data.scanned_barcode || data.barcode_number || '',
      password: data.password,
      jobId: data.jobId,
      createdAt: new Date()
    };
    
    // Log what's being inserted to database
    console.log("Inserting into database:", quarantineEntry);

    const result = await collection.insertOne(quarantineEntry);

    return res.json({ success: true, insertedId: result.insertedId });
})

app.get('/docuflow/citi_data/fetch-job-by-agreement', async (req, res) => {
  const agreementNumber = req.query.agreementNumber;
  const jobId = req.query.jobId;

  if (!agreementNumber) {
    return res.status(400).json({ error: 'Missing agreement number' });
  }

  console.log(`Fetching job for Agreement Number: ${agreementNumber}, Job ID hint: ${jobId}`);

  try {
    await client.connect();
    const db = client.db(DB_NAME);
    const collection = db.collection('jobs');

    const agreementNumAsString = String(agreementNumber).trim();
    const agreementNumAsInt = parseInt(agreementNumAsString, 10);
    
    console.log(`Looking for agreement ${agreementNumber} (${agreementNumAsInt})`);
    
    // Find the job that contains this agreement number
    // First try exact match by agreement number at the top level
    let job = await collection.findOne({ 'Agreement Number': { $in: [agreementNumAsString, agreementNumAsInt] } });
    
    // If job is found, return it
    if (job) {
      console.log(`Found job with Agreement Number at top level: ${job._id}`);
      return res.status(200).json(job);
    }
    
    // If not found, try looking for the agreement number inside the data array
    console.log('Not found at top level, checking in data arrays...');
    job = await collection.findOne({
      'data': {
        $elemMatch: {
          'Agreement Number': { $in: [agreementNumAsString, agreementNumAsInt] }
        }
      }
    });
    
    if (job) {
      console.log(`Found job with Agreement Number in data array: ${job._id}`);
      return res.status(200).json(job);
    }
    
    // As a fallback, if jobId is provided, try to find that specific job
    if (jobId) {
      console.log(`No job found with Agreement Number, trying with Job ID: ${jobId}`);
      
      let query;
      if (isValidObjectId(jobId)) {
        query = { _id: new ObjectId(jobId) };
      } else {
        query = { batchId: jobId };
      }
      
      job = await collection.findOne(query);
      
      if (job) {
        console.log(`Found job by ID: ${job._id}`);
        return res.status(200).json(job);
      }
    }
    
    // If we reach here, no job was found
    console.log(`No job found for Agreement Number: ${agreementNumber}`);
    return res.status(404).json({ error: 'No job found' });
    
  } catch (error) {
    console.error('Error fetching job by agreement:', error);
    return res.status(500).json({ 
      error: 'Failed to fetch job',
      details: error instanceof Error ? error.message : String(error)
    });
  } finally {
    if (client) {
      await client.close();
    }
  }
});

app.get('/docuflow/citi_data/get-citi-entry', async (req, res) => {
  const agreementNumber = req.query.agreementNumber;

  if (!process.env.MONGODB_URI) {
    return res.status(500).json({ error: 'MongoDB URI not found' });
  }

  let client;
  try {
    client = new MongoClient(process.env.MONGODB_URI);
    await client.connect();
    const db = client.db(DB_NAME);
    const collection = db.collection('citi_data');

    const citiEntry = await collection.find({
      'Agreement Number': agreementNumber
    }).toArray();

    return res.status(200).json(citiEntry);
  } catch (error) {
    console.error('Error fetching Citi entry:', error);
    return res.status(500).json({ error: 'Failed to fetch Citi entry' });
  } finally {
    if (client) {
      await client.close();
    }
  }
});

app.post("/docuflow/imageprocess", async (req, res) => {
    await client.connect();
    const db = client.db('ocr_database1');
    const collection = db.collection('agreements');

    const result = await collection.insertOne(req.body);

    const dbId = result.insertedId;
    const dbSaved = true;

    return res.json({
      structured_data: req.body,
      database_saved: dbSaved,
      database_id: dbId?.toString()
    });
})

app.get('/docuflow/jobs/fetch-job-details', async (req, res) => {
  const batchId = req.query.batchId;
  
  if (!batchId) {
    return res.status(400).json({ error: 'Missing batchId parameter' });
  }

  if (!process.env.MONGODB_URI) {
    return res.status(500).json({ error: 'MongoDB URI not found' });
  }

  let client;
  try {
    client = new MongoClient(process.env.MONGODB_URI);
    await client.connect();
    const db = client.db(DB_NAME);
    const collection = db.collection('jobs');
    
    // Find all documents for this batch ID
    const documents = await collection.find({ batchId }).toArray();
    
    // Format into the expected structure
    const job = {
      _id: batchId,
      batchId,
      data: documents,
      createdAt: documents[0]?.createdAt || new Date(),
      updatedAt: documents[0]?.updatedAt || new Date()
    };
    
    return res.status(200).json(job);
  } catch (error) {
    console.error('Error fetching job details:', error);
    const errorMessage = error instanceof Error ? error.message : 'Unknown error';
    return res.status(500).json({ error: 'Failed to fetch job details', details: errorMessage });
  } finally {
    if (client) {
      await client.close();
    }
  }
});

// Start the server
const PORT = process.env.PORT || 5672;
app.listen(PORT, () => {
  console.log(`🚀 Server running on http://localhost:${PORT}`);
});

function isValidObjectId(id) {
  try {
    new ObjectId(id);
    return true;
  } catch (error) {
    return false;
  }
}

// MongoDB connection details
const MONGODB_URI = process.env.MONGODB_URI || "mongodb://localhost:27017/";
const MONGODB_DBNAME = process.env.MONGODB_DBNAME || "DB_NAME";

// Helper functions for data validation
function checkCustomerNameMatch(row) {
  const reasons = [];
  if (row['Bank Name'] === 'Axis Bank' && row['Assign Customer Name'] && row['Matched Customer Name']) {
    const assignNames = row['Assign Customer Name'].trim().split(/\s+/);
    const matchedNames = row['Matched Customer Name'].trim().split(/\s+/);
    const matchFound = assignNames.some(assignPart => 
      matchedNames.some(matchedPart => 
        matchedPart.toLowerCase().includes(assignPart.toLowerCase())
      )
    );
    return matchFound ? '' : 'Not Okay (No partial match found)';
  } else if (row['Bank Name'] === 'Axis Bank') {
    if (!row['Assign Customer Name']) reasons.push("Assign Customer Name is blank");
    if (!row['Matched Customer Name']) reasons.push("Matched Customer Name is blank");
    return reasons.length ? `Not Okay (${reasons.join(', ')})` : '';
  }
  return '';
}

function checkAgreementNumber(row) {
  const reasons = [];
  const agreementNumber = row['Matched Agreement Number']?.toString();
  
  if (agreementNumber) {
    if (agreementNumber.length !== 10) {
      reasons.push(`Length is ${agreementNumber.length}, expected 10`);
    }
    if (!agreementNumber.startsWith('5')) {
      reasons.push("Does not start with '5'");
    }
    return reasons.length ? `Not Okay (${reasons.join(', ')})` : '';
  }
  return 'Not Okay (Blank)';
}

function checkDate(row) {
  const date = row['Matched Date'];
  if (date && typeof date === 'string') {
    const dateRegex = /^\d{2}-\d{2}-\d{4}$/;
    return dateRegex.test(date) ? '' : `Not Okay (Format not DD-MM-YYYY: ${date})`;
  }
  return 'Not Okay (Blank)';
}

function checkBarcode(row) {
  const reasons = [];
  const barcode = row['barcode_number']?.toString() || '';
  
  if (barcode) {
    if (barcode.startsWith('FJ')) {
      if (barcode.length !== 15) {
        reasons.push(`Starts with FJ but length is ${barcode.length}, expected 15`);
      }
    } else if (![13, 14, 15].includes(barcode.length)) {
      reasons.push(`Length is ${barcode.length}, not 13, 14, 15, or FJ15`);
    }
  } else {
    reasons.push("Blank");
  }
  return reasons.length ? `Not Okay (${reasons.join(', ')})` : '';
}

// Apply quality checks to data
function applyQualityChecks(data) {
  return data
    .filter(row => row.status && row.status !== '')
    .map(row => {
      const agreementCheck = checkAgreementNumber(row);
      const customerCheck = checkCustomerNameMatch(row);
      const dateCheck = checkDate(row);
      const barcodeCheck = checkBarcode(row);

      // Calculate confidence percentage
      const checks = [agreementCheck, customerCheck, dateCheck, barcodeCheck];
      const confidencePercentage = (checks.filter(check => check === '').length / 4) * 100;

      return {
        ...row,
        'Agreement Number Check': agreementCheck,
        'Customer Name Match': customerCheck,
        'Date Check': dateCheck,
        'Barcode Check': barcodeCheck,
        'Confidence Percentage': confidencePercentage
      };
    });
}

// Generate Excel file from data
function generateExcelFile(data, filename) {
  // Create workbook and worksheet
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);

  // Set column widths
  const colWidths = [
    { wch: 15 }, // Bank Name
    { wch: 10 }, // Confidence Percentage
    { wch: 20 }, // Assign Agreement Number
    { wch: 30 }, // Assign Customer Name
    { wch: 20 }, // Matched Agreement Number
    { wch: 30 }, // Matched Customer Name
    { wch: 15 }, // Matched Date
    { wch: 10 }, // status
    { wch: 15 }, // storageBoxNumber
    { wch: 15 }, // scanBoxNumber
    { wch: 15 }, // barcode_number
    { wch: 15 }, // scanned_barcode
    { wch: 15 }, // password
    { wch: 20 }, // jobId
    { wch: 20 }, // createdAt
    { wch: 30 }, // Agreement Number Check
    { wch: 30 }, // Customer Name Match
    { wch: 30 }, // Date Check
    { wch: 30 }, // Barcode Check
    { wch: 30 }  // scanned_barcode
  ];
  ws['!cols'] = colWidths;

  // Add the worksheet to workbook
  XLSX.utils.book_append_sheet(wb, ws, 'Combined Data');

  // Generate buffer
  return XLSX.write(wb, { 
    type: 'buffer', 
    bookType: 'xlsx',
    bookSST: false
  });
}

// Jobs export route - handles both GET and POST
app.route('/docuflow/jobs/export')
  .get(async (req, res) => {
    try {
      const jobId = req.query.jobId;
      console.log(`GET request for jobs export with jobId: ${jobId}`);
      
      const client = await MongoClient.connect(MONGODB_URI);
      const db = client.db(MONGODB_DBNAME);
      console.log(`Connected to database: ${MONGODB_DBNAME}`);
      
      let allData = [];
      
      // If jobId is provided, export only that job
      if (jobId) {
        console.log(`Searching for job with batchId: ${jobId}`);
        const job = await db.collection('jobs').findOne({ batchId: jobId.toString() });
        
        if (!job) {
          console.log(`Job not found with batchId: ${jobId}`);
          await client.close();
          return res.status(404).json({ error: 'Job not found' });
        }
        
        console.log(`Job found: ${JSON.stringify(job._id)}`);
        console.log(`Job data structure: ${job.data ? 'Has data array' : 'No data array'}`);
        
        // Transform the job data for Excel
        let enhancedData = [];
        
        if (job.data && Array.isArray(job.data)) {
          console.log(`Job has ${job.data.length} items in data array`);
          enhancedData = job.data.map((row, index) => {
            return {
              'Bank Name': 'Axis Bank',
              'Assign Agreement Number': row['Agreement Number'],
              'Assign Customer Name': job['Customer Name'],
              'Matched Agreement Number': row.matchedData?.agreement_number,
              'Matched Customer Name': row.matchedData?.customer_name,
              'Matched Date': row.matchedData?.date,
              'status': row.status || 'Pending',
              'storageBoxNumber': row.storageBoxNumber || 'N/A',
              'scanBoxNumber': row.scanBoxNumber || 'N/A',
              'barcode_number': row.barcode_number || 'N/A',
              'scanned_barcode': row.scanned_barcode || 'N/A',
              'password': null,
              'jobId': job.batchId,
              'createdAt': job.updatedAt
            };
          });
        } else {
          // Handle jobs documents without the 'data' array
          console.log(`Job has no data array, using direct properties`);
          enhancedData = [{
            'Bank Name': 'Axis Bank',
            'Assign Agreement Number': job['Agreement Number'],
            'Assign Customer Name': job['Customer Name'],
            'Matched Agreement Number': null,
            'Matched Customer Name': null,
            'Matched Date': null,
            'status': job.status || 'Pending',
            'storageBoxNumber': 'N/A',
            'scanBoxNumber': 'N/A',
            'barcode_number': 'N/A',
            'scanned_barcode': 'N/A',
            'password': null,
            'jobId': job.batchId,
            'createdAt': job.updatedAt
          }];
        }
        
        allData = enhancedData;
      } else {
        // Export all jobs data
        console.log('Exporting all jobs data');
        const jobsData = await db.collection('jobs').find().toArray();
        console.log(`Found ${jobsData.length} records in jobs`);
        
        jobsData.forEach(job => {
          if (job.data && Array.isArray(job.data)) {
            job.data.forEach(item => {
              allData.push({
                'Bank Name': 'Axis Bank',
                'Assign Agreement Number': item['Agreement Number'],
                'Assign Customer Name': job['Customer Name'],
                'Matched Agreement Number': item.matchedData?.agreement_number,
                'Matched Customer Name': item.matchedData?.customer_name,
                'Matched Date': item.matchedData?.date,
                'status': item.status,
                'storageBoxNumber': item.storageBoxNumber,
                'scanBoxNumber': item.scanBoxNumber,
                'barcode_number': item.barcode_number,
                'scanned_barcode': item.scanned_barcode,
                'password': null,
                'jobId': job.batchId,
                'createdAt': job.updatedAt
              });
            });
          } else {
            allData.push({
              'Bank Name': 'Axis Bank',
              'Assign Agreement Number': job['Agreement Number'],
              'Assign Customer Name': job['Customer Name'],
              'Matched Agreement Number': null,
              'Matched Customer Name': null,
              'Matched Date': null,
              'status': job.status,
              'storageBoxNumber': null,
              'scanBoxNumber': null,
              'barcode_number': null,
              'scanned_barcode': null,
              'password': null,
              'jobId': job.batchId,
              'createdAt': job.updatedAt
            });
          }
        });
      }
      
      // Apply quality checks
      const processedData = applyQualityChecks(allData);
      console.log(`Processed ${processedData.length} rows with quality checks`);
      
      // Generate Excel file
      const filename = jobId ? `job_${jobId}_details.xlsx` : 'jobs_export.xlsx';
      const excelBuffer = generateExcelFile(processedData, filename);
      
      // Return the Excel file
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', `attachment; filename=${filename}`);
      res.send(excelBuffer);
      
      await client.close();
    } catch (error) {
      console.error('Jobs export error:', error);
      res.status(500).json({ 
        error: 'Failed to generate jobs export',
        details: error.message || 'Unknown error'
      });
    }
  })
  .post(async (req, res) => {
    try {
      const jobId = req.body.jobId;
      console.log(`POST request for jobs export with jobId: ${jobId}`);
      
      // Reuse the same logic as GET but with body parameters
      const client = await MongoClient.connect(MONGODB_URI);
      const db = client.db(MONGODB_DBNAME);
      
      // Same logic as GET handler...
      // (Identical code to GET handler from here)
      let allData = [];
      
      if (jobId) {
        const job = await db.collection('jobs').findOne({ batchId: jobId.toString() });
        
        if (!job) {
          await client.close();
          return res.status(404).json({ error: 'Job not found' });
        }
        
        let enhancedData = [];
        
        if (job.data && Array.isArray(job.data)) {
          enhancedData = job.data.map((row) => {
            return {
              'Bank Name': 'Axis Bank',
              'Assign Agreement Number': row['Agreement Number'],
              'Assign Customer Name': job['Customer Name'],
              'Matched Agreement Number': row.matchedData?.agreement_number,
              'Matched Customer Name': row.matchedData?.customer_name,
              'Matched Date': row.matchedData?.date,
              'status': row.status || 'Pending',
              'storageBoxNumber': row.storageBoxNumber || 'N/A',
              'scanBoxNumber': row.scanBoxNumber || 'N/A',
              'barcode_number': row.barcode_number || 'N/A',
              'scanned_barcode': row.scanned_barcode || 'N/A',
              'password': null,
              'jobId': job.batchId,
              'createdAt': job.updatedAt
            };
          });
        } else {
          enhancedData = [{
            'Bank Name': 'Axis Bank',
            'Assign Agreement Number': job['Agreement Number'],
            'Assign Customer Name': job['Customer Name'],
            'Matched Agreement Number': null,
            'Matched Customer Name': null,
            'Matched Date': null,
            'status': job.status || 'Pending',
            'storageBoxNumber': 'N/A',
            'scanBoxNumber': 'N/A',
            'barcode_number': 'N/A',
            'scanned_barcode': 'N/A',
            'password': null,
            'jobId': job.batchId,
            'createdAt': job.updatedAt
          }];
        }
        
        allData = enhancedData;
      } else {
        const jobsData = await db.collection('jobs').find().toArray();
        
        jobsData.forEach(job => {
          if (job.data && Array.isArray(job.data)) {
            job.data.forEach(item => {
              allData.push({
                'Bank Name': 'Axis Bank',
                'Assign Agreement Number': item['Agreement Number'],
                'Assign Customer Name': job['Customer Name'],
                'Matched Agreement Number': item.matchedData?.agreement_number,
                'Matched Customer Name': item.matchedData?.customer_name,
                'Matched Date': item.matchedData?.date,
                'status': item.status,
                'storageBoxNumber': item.storageBoxNumber,
                'scanBoxNumber': item.scanBoxNumber,
                'barcode_number': item.barcode_number,
                'scanned_barcode': item.scanned_barcode,
                'password': null,
                'jobId': job.batchId,
                'createdAt': job.updatedAt
              });
            });
          } else {
            allData.push({
              'Bank Name': 'Axis Bank',
              'Assign Agreement Number': job['Agreement Number'],
              'Assign Customer Name': job['Customer Name'],
              'Matched Agreement Number': null,
              'Matched Customer Name': null,
              'Matched Date': null,
              'status': job.status,
              'storageBoxNumber': null,
              'scanBoxNumber': null,
              'barcode_number': null,
              'scanned_barcode': null,
              'password': null,
              'jobId': job.batchId,
              'createdAt': job.updatedAt
            });
          }
        });
      }
      
      const processedData = applyQualityChecks(allData);
      const filename = jobId ? `job_${jobId}_details.xlsx` : 'jobs_export.xlsx';
      const excelBuffer = generateExcelFile(processedData, filename);
      
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', `attachment; filename=${filename}`);
      res.send(excelBuffer);
      
      await client.close();
    } catch (error) {
      console.error('Jobs export error:', error);
      res.status(500).json({ 
        error: 'Failed to generate jobs export',
        details: error.message || 'Unknown error'
      });
    }
  });

// Citi-data export route - handles both GET and POST
app.route('/docuflow/citi-data/export')
  .get(async (req, res) => {
    try {
      console.log('GET request for citi-data export');
      
      const client = await MongoClient.connect(MONGODB_URI);
      const db = client.db(MONGODB_DBNAME);
      console.log(`Connected to database: ${MONGODB_DBNAME}`);
      
      // Fetch data from citi_data
      console.log('Fetching data from citi_data collection');
      const citiData = await db.collection('citi_data').find({}, { projection: { matchedImage: 0 } }).toArray();
      console.log(`Found ${citiData.length} records in citi_data`);
      
      let allData = [];
      
      citiData.forEach(doc => {
        allData.push({
          'Bank Name': 'Citi Bank',
          'Assign Agreement Number': null,
          'Assign Customer Name': null,
          'Matched Agreement Number': doc['Agreement Number'],
          'Matched Customer Name': doc['Customer Name'],
          'Matched Date': doc.date,
          'status': doc.status,
          'storageBoxNumber': doc.storageBoxNumber,
          'scanBoxNumber': doc.scanBoxNumber,
          'barcode_number': doc.barcode_number,
          'scanned_barcode': doc.scanned_barcode,
          'password': doc.password,
          'jobId': doc.jobId,
          'createdAt': doc.createdAt
        });
      });
      
      // Fetch data from quarantine_data
      console.log('Fetching data from quarantine_data collection');
      const quarantineData = await db.collection('quarantine_data').find({}, { projection: { matchedImage: 0 } }).toArray();
      console.log(`Found ${quarantineData.length} records in quarantine_data`);
      
      quarantineData.forEach(doc => {
        allData.push({
          'Bank Name': 'Quarantined',
          'Assign Agreement Number': null,
          'Assign Customer Name': null,
          'Matched Agreement Number': doc['Agreement Number'],
          'Matched Customer Name': doc['Customer Name'],
          'Matched Date': doc.date,
          'status': doc.status,
          'storageBoxNumber': doc.storageBoxNumber,
          'scanBoxNumber': doc.scanBoxNumber,
          'barcode_number': doc.barcode_number,
          'scanned_barcode': doc.scanned_barcode,
          'password': doc.password,
          'jobId': doc.jobId,
          'createdAt': doc.createdAt
        });
      });
      
      // Apply quality checks
      const processedData = applyQualityChecks(allData);
      console.log(`Processed ${processedData.length} rows with quality checks`);
      
      // Generate Excel file
      const excelBuffer = generateExcelFile(processedData, 'citi_data_export.xlsx');
      
      // Return the Excel file
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename=citi_data_export.xlsx');
      res.send(excelBuffer);
      
      await client.close();
    } catch (error) {
      console.error('Citi-data export error:', error);
      res.status(500).json({ 
        error: 'Failed to generate citi-data export',
        details: error.message || 'Unknown error'
      });
    }
  })
  .post(async (req, res) => {
    try {
      console.log('POST request for citi-data export');
      
      // Reuse the same logic as GET
      const client = await MongoClient.connect(MONGODB_URI);
      const db = client.db(MONGODB_DBNAME);
      
      // Same logic as GET handler...
      // (Identical code to GET handler from here)
      const citiData = await db.collection('citi_data').find({}, { projection: { matchedImage: 0 } }).toArray();
      
      let allData = [];
      
      citiData.forEach(doc => {
        allData.push({
          'Bank Name': 'Citi Bank',
          'Assign Agreement Number': null,
          'Assign Customer Name': null,
          'Matched Agreement Number': doc['Agreement Number'],
          'Matched Customer Name': doc['Customer Name'],
          'Matched Date': doc.date,
          'status': doc.status,
          'storageBoxNumber': doc.storageBoxNumber,
          'scanBoxNumber': doc.scanBoxNumber,
          'barcode_number': doc.barcode_number,
          'scanned_barcode': doc.scanned_barcode,
          'password': doc.password,
          'jobId': doc.jobId,
          'createdAt': doc.createdAt
        });
      });
      
      const quarantineData = await db.collection('quarantine_data').find({}, { projection: { matchedImage: 0 } }).toArray();
      
      quarantineData.forEach(doc => {
        allData.push({
          'Bank Name': 'Quarantined',
          'Assign Agreement Number': null,
          'Assign Customer Name': null,
          'Matched Agreement Number': doc['Agreement Number'],
          'Matched Customer Name': doc['Customer Name'],
          'Matched Date': doc.date,
          'status': doc.status,
          'storageBoxNumber': doc.storageBoxNumber,
          'scanBoxNumber': doc.scanBoxNumber,
          'barcode_number': doc.barcode_number,
          'scanned_barcode': doc.scanned_barcode,
          'password': doc.password,
          'jobId': doc.jobId,
          'createdAt': doc.createdAt
        });
      });
      
      const processedData = applyQualityChecks(allData);
      const excelBuffer = generateExcelFile(processedData, 'citi_data_export.xlsx');
      
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename=citi_data_export.xlsx');
      res.send(excelBuffer);
      
      await client.close();
    } catch (error) {
      console.error('Citi-data export error:', error);
      res.status(500).json({ 
        error: 'Failed to generate citi-data export',
        details: error.message || 'Unknown error'
      });
    }
  });
