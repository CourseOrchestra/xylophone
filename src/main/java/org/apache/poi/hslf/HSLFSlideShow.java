/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

package org.apache.poi.hslf;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Hashtable;
import java.util.List;
import java.util.Map;
import java.util.NavigableMap;
import java.util.TreeMap;

import org.apache.poi.POIDocument;
import org.apache.poi.hslf.exceptions.CorruptPowerPointFileException;
import org.apache.poi.hslf.exceptions.EncryptedPowerPointFileException;
import org.apache.poi.hslf.exceptions.HSLFException;
import org.apache.poi.hslf.record.CurrentUserAtom;
import org.apache.poi.hslf.record.ExOleObjStg;
import org.apache.poi.hslf.record.PersistPtrHolder;
import org.apache.poi.hslf.record.PersistRecord;
import org.apache.poi.hslf.record.PositionDependentRecord;
import org.apache.poi.hslf.record.Record;
import org.apache.poi.hslf.record.RecordTypes;
import org.apache.poi.hslf.record.UserEditAtom;
import org.apache.poi.hslf.usermodel.ObjectData;
import org.apache.poi.hslf.usermodel.PictureData;
import org.apache.poi.poifs.filesystem.DirectoryNode;
import org.apache.poi.poifs.filesystem.DocumentEntry;
import org.apache.poi.poifs.filesystem.DocumentInputStream;
import org.apache.poi.poifs.filesystem.EntryUtils;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.util.LittleEndian;
import org.apache.poi.util.POILogFactory;
import org.apache.poi.util.POILogger;

/**
 * This class contains the main functionality for the Powerpoint file
 * "reader". It is only a very basic class for now
 *
 * @author Nick Burch
 */
public final class HSLFSlideShow extends POIDocument {
    public static final int UNSET_OFFSET = -1;
    
    // For logging
    private POILogger logger = POILogFactory.getLogger(this.getClass());

	// Holds metadata on where things are in our document
	private CurrentUserAtom currentUser;

	// Low level contents of the file
	private byte[] _docstream;

	// Low level contents
	private Record[] _records;

	// Raw Pictures contained in the pictures stream
	private List<PictureData> _pictures;

    // Embedded objects stored in storage records in the document stream, lazily populated.
    private ObjectData[] _objects;
    
    /**
	 * Returns the underlying POIFSFileSystem for the document
	 *  that is open.
	 */
	protected POIFSFileSystem getPOIFSFileSystem() {
		return directory.getFileSystem();
	}

   /**
    * Returns the directory in the underlying POIFSFileSystem for the 
    *  document that is open.
    */
   protected DirectoryNode getPOIFSDirectory() {
      return directory;
   }

	/**
	 * Constructs a Powerpoint document from fileName. Parses the document
	 * and places all the important stuff into data structures.
	 *
	 * @param fileName The name of the file to read.
	 * @throws IOException if there is a problem while parsing the document.
	 */
	public HSLFSlideShow(String fileName) throws IOException
	{
		this(new FileInputStream(fileName));
	}

	/**
	 * Constructs a Powerpoint document from an input stream. Parses the
	 * document and places all the important stuff into data structures.
	 *
	 * @param inputStream the source of the data
	 * @throws IOException if there is a problem while parsing the document.
	 */
	public HSLFSlideShow(InputStream inputStream) throws IOException {
		//do Ole stuff
		this(new POIFSFileSystem(inputStream));
	}

	/**
	 * Constructs a Powerpoint document from a POIFS Filesystem. Parses the
	 * document and places all the important stuff into data structures.
	 *
	 * @param filesystem the POIFS FileSystem to read from
	 * @throws IOException if there is a problem while parsing the document.
	 */
	public HSLFSlideShow(POIFSFileSystem filesystem) throws IOException
	{
		this(filesystem.getRoot());
	}

   /**
    * Constructs a Powerpoint document from a POIFS Filesystem. Parses the
    * document and places all the important stuff into data structures.
    *
    * @param filesystem the POIFS FileSystem to read from
    * @throws IOException if there is a problem while parsing the document.
    */
   public HSLFSlideShow(NPOIFSFileSystem filesystem) throws IOException
   {
      this(filesystem.getRoot());
   }

   /**
    * Constructs a Powerpoint document from a specific point in a
    *  POIFS Filesystem. Parses the document and places all the
    *  important stuff into data structures.
    *
    * @deprecated Use {@link #HSLFSlideShow(DirectoryNode)} instead
    * @param dir the POIFS directory to read from
    * @param filesystem the POIFS FileSystem to read from
    * @throws IOException if there is a problem while parsing the document.
    */
	@Deprecated
   public HSLFSlideShow(DirectoryNode dir, POIFSFileSystem filesystem) throws IOException
   {
      this(dir);
   }
   
	/**
	 * Constructs a Powerpoint document from a specific point in a
	 *  POIFS Filesystem. Parses the document and places all the
	 *  important stuff into data structures.
	 *
	 * @param dir the POIFS directory to read from
	 * @throws IOException if there is a problem while parsing the document.
	 */
	public HSLFSlideShow(DirectoryNode dir) throws IOException
	{
		super(dir);

		// First up, grab the "Current User" stream
		// We need this before we can detect Encrypted Documents
		readCurrentUserStream();

		// Next up, grab the data that makes up the
		//  PowerPoint stream
		readPowerPointStream();

		// Check to see if we have an encrypted document,
		//  bailing out if we do
		boolean encrypted = EncryptedSlideShow.checkIfEncrypted(this);
		if(encrypted) {
			throw new EncryptedPowerPointFileException("Encrypted PowerPoint files are not supported");
		}

		// Now, build records based on the PowerPoint stream
		buildRecords();

		// Look for any other streams
		readOtherStreams();
	}
	
	
	
	/**
	 * Constructs a new, empty, Powerpoint document.
	 */
	public static final HSLFSlideShow create() {
		InputStream is = HSLFSlideShow.class.getResourceAsStream("data/empty.ppt");
		if (is == null) {
			throw new RuntimeException("Missing resource 'empty.ppt'");
		}
		try {
			return new HSLFSlideShow(is);
		} catch (IOException e) {
			throw new RuntimeException(e);
		}
	}

	/**
	 * Extracts the main PowerPoint document stream from the
	 *  POI file, ready to be passed
	 *
	 * @throws IOException
	 */
	private void readPowerPointStream() throws IOException
	{
		// Get the main document stream
		DocumentEntry docProps =
			(DocumentEntry)directory.getEntry("PowerPoint Document");

		// Grab the document stream
		_docstream = new byte[docProps.getSize()];
		directory.createDocumentInputStream("PowerPoint Document").read(_docstream);
	}

	/**
	 * Builds the list of records, based on the contents
	 *  of the PowerPoint stream
	 */
	private void buildRecords()
	{
		// The format of records in a powerpoint file are:
		//   <little endian 2 byte "info">
		//   <little endian 2 byte "type">
		//   <little endian 4 byte "length">
		// If it has a zero length, following it will be another record
		//		<xx xx yy yy 00 00 00 00> <xx xx yy yy zz zz zz zz>
		// If it has a length, depending on its type it may have children or data
		// If it has children, these will follow straight away
		//		<xx xx yy yy zz zz zz zz <xx xx yy yy zz zz zz zz>>
		// If it has data, this will come straigh after, and run for the length
		//      <xx xx yy yy zz zz zz zz dd dd dd dd dd dd dd>
		// All lengths given exclude the 8 byte record header
		// (Data records are known as Atoms)

		// Document should start with:
		//   0F 00 E8 03 ## ## ## ##
	    //     (type 1000 = document, info 00 0f is normal, rest is document length)
		//   01 00 E9 03 28 00 00 00
		//     (type 1001 = document atom, info 00 01 normal, 28 bytes long)
		//   80 16 00 00 E0 10 00 00 xx xx xx xx xx xx xx xx
		//   05 00 00 00 0A 00 00 00 xx xx xx
		//     (the contents of the document atom, not sure what it means yet)
		//   (records then follow)

		// When parsing a document, look to see if you know about that type
		//  of the current record. If you know it's a type that has children,
		//  process the record's data area looking for more records
		// If you know about the type and it doesn't have children, either do
		//  something with the data (eg TextRun) or skip over it
		// If you don't know about the type, play safe and skip over it (using
		//  its length to know where the next record will start)
		//

        _records = read(_docstream, (int)currentUser.getCurrentEditOffset());
	}

	private Record[] read(byte[] docstream, int usrOffset){
        //sort found records by offset.
        //(it is not necessary but SlideShow.findMostRecentCoreRecords() expects them sorted)
	    NavigableMap<Integer,Record> records = new TreeMap<Integer,Record>(); // offset -> record
        Map<Integer,Integer> persistIds = new HashMap<Integer,Integer>(); // offset -> persistId
        initRecordOffsets(docstream, usrOffset, records, persistIds);
        
        for (Map.Entry<Integer,Record> entry : records.entrySet()) {
            Integer offset = entry.getKey();
            Record record = entry.getValue();
            Integer persistId = persistIds.get(offset);
            if (record == null) {
                // all plain records have been already added,
                // only new records need to be decrypted (tbd #35897)
                record = Record.buildRecordAtOffset(docstream, offset);
                entry.setValue(record);
            }
            
            if (record instanceof PersistRecord) {
                ((PersistRecord)record).setPersistId(persistId);
            }            
        }
        
        return records.values().toArray(new Record[records.size()]);
    }

    private void initRecordOffsets(byte[] docstream, int usrOffset, NavigableMap<Integer,Record> recordMap, Map<Integer,Integer> offset2id) {
        while (usrOffset != 0){
            UserEditAtom usr = (UserEditAtom) Record.buildRecordAtOffset(docstream, usrOffset);
            recordMap.put(usrOffset, usr);
            
            int psrOffset = usr.getPersistPointersOffset();
            PersistPtrHolder ptr = (PersistPtrHolder)Record.buildRecordAtOffset(docstream, psrOffset);
            recordMap.put(psrOffset, ptr);
            
            for(Map.Entry<Integer,Integer> entry : ptr.getSlideLocationsLookup().entrySet()) {
                Integer offset = entry.getValue();
                Integer id = entry.getKey();
                recordMap.put(offset, null); // reserve a slot for the record
                offset2id.put(offset, id);
            }
            
            usrOffset = usr.getLastUserEditAtomOffset();

            // check for corrupted user edit atom and try to repair it
            // if the next user edit atom offset is already known, we would go into an endless loop
            if (usrOffset > 0 && recordMap.containsKey(usrOffset)) {
                // a user edit atom is usually located 36 byte before the smallest known record offset 
                usrOffset = recordMap.firstKey()-36;
                // check that we really are located on a user edit atom
                int ver_inst = LittleEndian.getUShort(docstream, usrOffset);
                int type = LittleEndian.getUShort(docstream, usrOffset+2);
                int len = LittleEndian.getInt(docstream, usrOffset+4);
                if (ver_inst == 0 && type == 4085 && (len == 0x1C || len == 0x20)) {
                    logger.log(POILogger.WARN, "Repairing invalid user edit atom");
                    usr.setLastUserEditAtomOffset(usrOffset);
                } else {
                    throw new CorruptPowerPointFileException("Powerpoint document contains invalid user edit atom");
                }
            }
        }       
    }

	/**
	 * Find the "Current User" stream, and load it
	 */
	private void readCurrentUserStream() {
		try {
			currentUser = new CurrentUserAtom(directory);
		} catch(IOException ie) {
			logger.log(POILogger.ERROR, "Error finding Current User Atom:\n" + ie);
			currentUser = new CurrentUserAtom();
		}
	}

	/**
	 * Find any other streams from the filesystem, and load them
	 */
	private void readOtherStreams() {
		// Currently, there aren't any
	}
	/**
	 * Find and read in pictures contained in this presentation.
	 * This is lazily called as and when we want to touch pictures.
	 */
    @SuppressWarnings("unused")
	private void readPictures() throws IOException {
        _pictures = new ArrayList<PictureData>();

        // if the presentation doesn't contain pictures - will use a null set instead
        if (!directory.hasEntry("Pictures")) return;
        
		DocumentEntry entry = (DocumentEntry)directory.getEntry("Pictures");
		byte[] pictstream = new byte[entry.getSize()];
		DocumentInputStream is = directory.createDocumentInputStream(entry);
		is.read(pictstream);
		is.close();

		
        int pos = 0;
		// An empty picture record (length 0) will take up 8 bytes
        while (pos <= (pictstream.length-8)) {
            int offset = pos;
            
            // Image signature
            int signature = LittleEndian.getUShort(pictstream, pos);
            pos += LittleEndian.SHORT_SIZE;
            // Image type + 0xF018
            int type = LittleEndian.getUShort(pictstream, pos);
            pos += LittleEndian.SHORT_SIZE;
            // Image size (excluding the 8 byte header)
            int imgsize = LittleEndian.getInt(pictstream, pos);
            pos += LittleEndian.INT_SIZE;

            // When parsing the BStoreDelay stream, [MS-ODRAW] says that we
            //  should terminate if the type isn't 0xf007 or 0xf018->0xf117
            if (!((type == 0xf007) || (type >= 0xf018 && type <= 0xf117)))
                break;

			// The image size must be 0 or greater
			// (0 is allowed, but odd, since we do wind on by the header each
			//  time, so we won't get stuck)
			if(imgsize < 0) {
				throw new CorruptPowerPointFileException("The file contains a picture, at position " + _pictures.size() + ", which has a negatively sized data length, so we can't trust any of the picture data");
			}

			// If they type (including the bonus 0xF018) is 0, skip it
			if(type == 0) {
				logger.log(POILogger.ERROR, "Problem reading picture: Invalid image type 0, on picture with length " + imgsize + ".\nYou document will probably become corrupted if you save it!");
				logger.log(POILogger.ERROR, "" + pos);
			} else {
				// Build the PictureData object from the data
				try {
					PictureData pict = PictureData.create(type - 0xF018);

                    // Copy the data, ready to pass to PictureData
                    byte[] imgdata = new byte[imgsize];
                    System.arraycopy(pictstream, pos, imgdata, 0, imgdata.length);
                    pict.setRawData(imgdata);

                    pict.setOffset(offset);
					_pictures.add(pict);
				} catch(IllegalArgumentException e) {
					logger.log(POILogger.ERROR, "Problem reading picture: " + e + "\nYou document will probably become corrupted if you save it!");
				}
			}

            pos += imgsize;
        }
	}

	/**
     * This is a helper functions, which is needed for adding new position dependent records
     * or finally write the slideshow to a file.
	 *
	 * @param os the stream to write to, if null only the references are updated
	 * @param interestingRecords a map of interesting records (PersistPtrHolder and UserEditAtom)
	 *        referenced by their RecordType. Only the very last of each type will be saved to the map.
	 *        May be null, if not needed. 
	 * @throws IOException
	 */
	public void updateAndWriteDependantRecords(OutputStream os, Map<RecordTypes.Type,PositionDependentRecord> interestingRecords)
	throws IOException {
        // For position dependent records, hold where they were and now are
        // As we go along, update, and hand over, to any Position Dependent
        //  records we happen across
        Hashtable<Integer,Integer> oldToNewPositions = new Hashtable<Integer,Integer>();

        // First pass - figure out where all the position dependent
        //   records are going to end up, in the new scheme
        // (Annoyingly, some powerpoint files have PersistPtrHolders
        //  that reference slides after the PersistPtrHolder)
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        for (Record record : _records) {
            if(record instanceof PositionDependentRecord) {
                PositionDependentRecord pdr = (PositionDependentRecord)record;
                int oldPos = pdr.getLastOnDiskOffset();
                int newPos = baos.size();
                pdr.setLastOnDiskOffset(newPos);
                if (oldPos != UNSET_OFFSET) {
                    // new records don't need a mapping, as they aren't in a relation yet
                    oldToNewPositions.put(Integer.valueOf(oldPos),Integer.valueOf(newPos));
                }
            }
            
            // Dummy write out, so the position winds on properly
            record.writeOut(baos);
        }
        baos = null;
        
        // For now, we're only handling PositionDependentRecord's that
        // happen at the top level.
        // In future, we'll need the handle them everywhere, but that's
        // a bit trickier
	    UserEditAtom usr = null;
        for (Record record : _records) {
            if (record instanceof PositionDependentRecord) {
                // We've already figured out their new location, and
                // told them that
                // Tell them of the positions of the other records though
                PositionDependentRecord pdr = (PositionDependentRecord)record;
                pdr.updateOtherRecordReferences(oldToNewPositions);
    
                // Grab interesting records as they come past
                // this will only save the very last record of each type
                RecordTypes.Type saveme = null;
                int recordType = (int)record.getRecordType();
                if (recordType == RecordTypes.PersistPtrIncrementalBlock.typeID) {
                    saveme = RecordTypes.PersistPtrIncrementalBlock;
                } else if (recordType == RecordTypes.UserEditAtom.typeID) {
                    saveme = RecordTypes.UserEditAtom;
                    usr = (UserEditAtom)pdr;
                }
                if (interestingRecords != null && saveme != null) {
                    interestingRecords.put(saveme,pdr);
                }
            }
            
            // Whatever happens, write out that record tree
            if (os != null) {
                record.writeOut(os);
            }
        }

        // Update and write out the Current User atom
        int oldLastUserEditAtomPos = (int)currentUser.getCurrentEditOffset();
        Integer newLastUserEditAtomPos = oldToNewPositions.get(oldLastUserEditAtomPos);
        if(usr == null || newLastUserEditAtomPos == null || usr.getLastOnDiskOffset() != newLastUserEditAtomPos) {
            throw new HSLFException("Couldn't find the new location of the last UserEditAtom that used to be at " + oldLastUserEditAtomPos);
        }
        currentUser.setCurrentEditOffset(usr.getLastOnDiskOffset());
	}
	
    /**
     * Writes out the slideshow file the is represented by an instance
     *  of this class.
     * It will write out the common OLE2 streams. If you require all
     *  streams to be written out, pass in preserveNodes
     * @param out The OutputStream to write to.
     * @throws IOException If there is an unexpected IOException from
     *           the passed in OutputStream
     */
    public void write(OutputStream out) throws IOException {
        // Write out, but only the common streams
        write(out,false);
    }
    /**
     * Writes out the slideshow file the is represented by an instance
     *  of this class.
     * If you require all streams to be written out (eg Marcos, embeded
     *  documents), then set preserveNodes to true
     * @param out The OutputStream to write to.
     * @param preserveNodes Should all OLE2 streams be written back out, or only the common ones?
     * @throws IOException If there is an unexpected IOException from
     *           the passed in OutputStream
     */
    public void write(OutputStream out, boolean preserveNodes) throws IOException {
        // Get a new Filesystem to write into
        POIFSFileSystem outFS = new POIFSFileSystem();

        // The list of entries we've written out
        List<String> writtenEntries = new ArrayList<String>(1);

        // Write out the Property Streams
        writeProperties(outFS, writtenEntries);

        ByteArrayOutputStream baos = new ByteArrayOutputStream();

        // For position dependent records, hold where they were and now are
        // As we go along, update, and hand over, to any Position Dependent
        // records we happen across
        updateAndWriteDependantRecords(baos, null);

        // Update our cached copy of the bytes that make up the PPT stream
        _docstream = baos.toByteArray();

        // Write the PPT stream into the POIFS layer
        ByteArrayInputStream bais = new ByteArrayInputStream(baos.toByteArray());
        outFS.createDocument(bais,"PowerPoint Document");
        writtenEntries.add("PowerPoint Document");
        currentUser.writeToFS(outFS);
        writtenEntries.add("Current User");


        // Write any pictures, into another stream
        if(_pictures == null) {
           readPictures();
        }
        if (_pictures.size() > 0) {
            ByteArrayOutputStream pict = new ByteArrayOutputStream();
            for (PictureData p : _pictures) {
                p.write(pict);
            }
            outFS.createDocument(
                new ByteArrayInputStream(pict.toByteArray()), "Pictures"
            );
            writtenEntries.add("Pictures");
        }

        // If requested, write out any other streams we spot
        if(preserveNodes) {
            EntryUtils.copyNodes(directory.getFileSystem(), outFS, writtenEntries);
        }

        // Send the POIFSFileSystem object out to the underlying stream
        outFS.writeFilesystem(out);
    }


	/* ******************* adding methods follow ********************* */

	/**
	 * Adds a new root level record, at the end, but before the last
	 *  PersistPtrIncrementalBlock.
	 */
	public synchronized int appendRootLevelRecord(Record newRecord) {
		int addedAt = -1;
		Record[] r = new Record[_records.length+1];
		boolean added = false;
		for(int i=(_records.length-1); i>=0; i--) {
			if(added) {
				// Just copy over
				r[i] = _records[i];
			} else {
				r[(i+1)] = _records[i];
				if(_records[i] instanceof PersistPtrHolder) {
					r[i] = newRecord;
					added = true;
					addedAt = i;
				}
			}
		}
		_records = r;
		return addedAt;
	}

	/**
	 * Add a new picture to this presentation.
     *
     * @return offset of this picture in the Pictures stream
	 */
	public int addPicture(PictureData img) {
	   // Process any existing pictures if we haven't yet
	   if(_pictures == null) {
         try {
            readPictures();
         } catch(IOException e) {
            throw new CorruptPowerPointFileException(e.getMessage());
         }
	   }
	   
	   // Add the new picture in
      int offset = 0;
	   if(_pictures.size() > 0) {
	      PictureData prev = _pictures.get(_pictures.size() - 1);
	      offset = prev.getOffset() + prev.getRawData().length + 8;
	   }
	   img.setOffset(offset);
	   _pictures.add(img);
	   return offset;
   }

	/* ******************* fetching methods follow ********************* */


	/**
	 * Returns an array of all the records found in the slideshow
	 */
	public Record[] getRecords() { return _records; }

	/**
	 * Returns an array of the bytes of the file. Only correct after a
	 *  call to open or write - at all other times might be wrong!
	 */
	public byte[] getUnderlyingBytes() { return _docstream; }

	/**
	 * Fetch the Current User Atom of the document
	 */
	public CurrentUserAtom getCurrentUserAtom() { return currentUser; }

	/**
	 *  Return array of pictures contained in this presentation
	 *
	 *  @return array with the read pictures or <code>null</code> if the
	 *  presentation doesn't contain pictures.
	 */
	public PictureData[] getPictures() {
	   if(_pictures == null) {
	      try {
	         readPictures();
	      } catch(IOException e) {
	         throw new CorruptPowerPointFileException(e.getMessage());
	      }
	   }
	   
		return _pictures.toArray(new PictureData[_pictures.size()]);
	}

    /**
     * Gets embedded object data from the slide show.
     *
     * @return the embedded objects.
     */
    public ObjectData[] getEmbeddedObjects() {
        if (_objects == null) {
            List<ObjectData> objects = new ArrayList<ObjectData>();
            for (Record r : _records) {
                if (r instanceof ExOleObjStg) {
                    objects.add(new ObjectData((ExOleObjStg)r));
                }
            }
            _objects = objects.toArray(new ObjectData[objects.size()]);
        }
        return _objects;
    }
}
