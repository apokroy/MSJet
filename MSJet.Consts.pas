unit MSJet.Consts;

interface

const
  /// <summary>
  /// This parameter indicates the relative or absolute file system path of the
  /// folder that will contain the checkpoint file for the instance. The path
  /// must be terminated with a backslash character, which indicates that the
  /// target path is a folder.
  /// </summary>
  jet_paramSystemPath = 0;

  /// <summary>
  /// This parameter indicates the relative or absolute file system path of
  /// the folder or file that will contain the temporary database for the instance.
  /// If the path is to a folder that will contain the temporary database then it
  /// must be terminated with a backslash character.
  /// </summary>
  jet_paramTempPath = 1;

  /// <summary>
  /// This parameter indicates the relative or absolute file system path of the
  /// folder that will contain the transaction logs for the instance. The path must
  /// be terminated with a backslash character, which indicates that the target path
  /// is a folder.
  /// </summary>
  jet_paramLogFilePath = 2;

  /// <summary>
  /// This parameter sets the three letter prefix used for many of the files used by
  /// the database engine. For example, the checkpoint file is called EDB.CHK by
  /// default because EDB is the default base name.
  /// </summary>
  jet_paramBaseName = 3;

  /// <summary>
  /// This parameter supplies an application specific string that will be added to
  /// any event log messages that are emitted by the database engine. This allows
  /// easy correlation of event log messages with the source application. By default
  /// the host application executable name will be used.
  /// </summary>
  jet_paramEventSource = 4;

  /// <summary>
  /// This parameter reserves the requested number of session resources for use by an
  /// instance. A session resource directly corresponds to a JET_SESID data type.
  /// This setting will affect how many sessions can be used at the same time.
  /// </summary>
  jet_paramMaxSessions = 5;

  /// <summary>
  /// This parameter reserves the requested number of B+ Tree resources for use by
  /// an instance. This setting will affect how many tables can be used at the same time.
  /// </summary>
  jet_paramMaxOpenTables = 6;

  /// <summary>
  /// This parameter reserves the requested number of cursor resources for use by an
  /// instance. A cursor resource directly corresponds to a JET_TABLEID data type.
  /// This setting will affect how many cursors can be used at the same time. A cursor
  /// resource cannot be shared by different sessions so this parameter must be set to
  /// a large enough value so that each session can use as many cursors as are required.
  /// </summary>
  jet_paramMaxCursors = 8;

  /// <summary>
  /// This parameter reserves the requested number of version store pages for use by an instance.
  /// </summary>
  jet_paramMaxVerPages = 9;

  /// <summary>
  /// This parameter reserves the requested number of temporary table resources for use
  /// by an instance. This setting will affect how many temporary tables can be used at
  /// the same time. If this system parameter is set to zero then no temporary database
  /// will be created and any activity that requires use of the temporary database will
  /// fail. This setting can be useful to avoid the I/O required to create the temporary
  /// database if it is known that it will not be used.
  /// </summary>
  /// <remarks>
  /// The use of a temporary table also requires a cursor resource.
  /// </remarks>
  jet_paramMaxTemporaryTables = 10;

  /// <summary>
  /// This parameter will configure the size of the transaction log files. Each
  /// transaction log file is a fixed size. The size is equal to the setting of
  /// this system parameter in units of 1024 bytes.
  /// </summary>
  jet_paramLogFileSize = 11;

  /// <summary>
  /// This parameter will configure the amount of memory used to cache log records
  /// before they are written to the transaction log file. The unit for this
  /// parameter is the sector size of the volume that holds the transaction log files.
  /// The sector size is almost always 512 bytes, so it is safe to assume that size
  /// for the unit. This parameter has an impact on performance. When the database
  /// engine is under heavy update load, this buffer can become full very rapidly.
  /// A larger cache size for the transaction log file is critical for good update
  /// performance under such a high load condition. The default is known to be too small
  /// for this case.
  /// Do not set this parameter to a number of buffers that is larger (in bytes) than
  /// half the size of a transaction log file.
  /// </summary>
  jet_paramLogBuffers = 12;

  /// <summary>
  /// This parameter configures how transaction log files are managed by the database
  /// engine. When circular logging is off, all transaction log files that are generated
  /// are retained on disk until they are no longer needed because a full backup of the
  /// database has been performed. When circular logging is on, only transaction log files
  /// that are younger than the current checkpoint are retained on disk. The benefit of
  /// this mode is that backups are not required to retire old transaction log files.
  /// </summary>
  jet_paramCircularLog = 17;

  /// <summary>
  /// This parameter controls the amount of space that is added to a database file each
  /// time it needs to grow to accommodate more data. The size is in database pages.
  /// </summary>
  jet_paramDbExtensionSize = 18;

  /// <summary>
  /// This parameter controls the initial size of the temporary database. The size is in
  /// database pages. A size of zero indicates that the default size of an ordinary
  /// database should be used. It is often desirable for small applications to configure
  /// the temporary database to be as small as possible. Setting this parameter to
  /// SystemParameters.PageTempDBSmallest will achieve the smallest temporary database possible.
  /// </summary>
  jet_paramPageTempDBMin = 19;

  /// <summary>
  /// This parameter configures the maximum size of the database page cache. The size
  /// is in database pages. If this parameter is left to its default value, then the
  /// maximum size of the cache will be set to the size of physical memory when JetInit
  /// is called.
  /// </summary>
  jet_paramCacheSizeMax = 23;

  /// <summary>
  /// This parameter controls how aggressively database pages are flushed from the
  /// database page cache to minimize the amount of time it will take to recover from a
  /// crash. The parameter is a threshold in bytes for about how many transaction log
  /// files will need to be replayed after a crash. If circular logging is enabled using
  /// JET_param.CircularLog then this parameter will also control the approximate amount
  /// of transaction log files that will be retained on disk.
  /// </summary>
  jet_paramCheckpointDepthMax = 24;

  /// <summary>
  /// This parameter controls the correlation interval of ESE's LRU-K page replacement
  /// algorithm.
  /// </summary>
  jet_paramLrukCorrInterval = 25;

  /// <summary>
  /// This parameter controls the timeout interval of ESE's LRU-K page replacement
  /// algorithm.
  /// </summary>
  jet_paramLrukTimeout = 28;

  /// <summary>
  /// This parameter controls how many database file I/Os can be queued
  /// per-disk in the host operating system at one time.  A larger value
  /// for this parameter can significantly help the performance of a large
  /// database application.
  /// </summary>
  jet_paramOutstandingIOMax = 30;

  /// <summary>
  /// This parameter controls when the database page cache begins evicting pages from the
  /// cache to make room for pages that are not cached. When the number of page buffers in the cache
  /// drops below this threshold then a background process will be started to replenish that pool
  /// of available buffers. This threshold is always relative to the maximum cache size as set by
  /// JET_paramCacheSizeMax. This threshold must also always be less than the stop threshold as
  /// set by JET_paramStopFlushThreshold.
  /// <para>
  /// The distance height of the start threshold will determine the response time that the database
  ///  page cache must have to produce available buffers before the application needs them. A high
  /// start threshold will give the background process more time to react. However, a high start
  /// threshold implies a higher stop threshold and that will reduce the effective size of the
  /// database page cache for modified pages (Windows 2000) or for all pages (Windows XP and later).
  /// </para>
  /// </summary>
  jet_paramStartFlushThreshold = 31;

  /// <summary>
  /// This parameter controls when the database page cache ends evicting pages from the cache to make
  /// room for pages that are not cached. When the number of page buffers in the cache rises above
  /// this threshold then the background process that was started to replenish that pool of available
  /// buffers is stopped. This threshold is always relative to the maximum cache size as set by
  /// JET_paramCacheSizeMax. This threshold must also always be greater than the start threshold
  /// as set by JET_paramStartFlushThreshold.
  /// <para>
  /// The distance between the start threshold and the stop threshold affects the efficiency with
  /// which database pages are flushed by the background process. A larger gap will make it
  /// more likely that writes to neighboring pages may be combined. However, a high stop
  /// threshold will reduce the effective size of the database page cache for modified
  /// pages (Windows 2000) or for all pages (Windows XP and later).
  /// </para>
  /// </summary>
  jet_paramStopFlushThreshold = 32;

  /// <summary>
  /// This parameter is the master switch that controls crash recovery for an instance.
  /// If this parameter is set to "On" then ARIES style recovery will be used to bring all
  /// databases in the instance to a consistent state in the event of a process or machine
  /// crash. If this parameter is set to "Off" then all databases in the instance will be
  /// managed without the benefit of crash recovery. That is to say, that if the instance
  /// is not shut down cleanly using JetTerm prior to the process exiting or machine shutdown
  /// then the contents of all databases in that instance will be corrupted.
  /// </summary>
  jet_paramRecovery = 34;

  /// <summary>
  /// This parameter controls the behavior of online defragmentation when initiated using
  /// </summary>
  jet_paramEnableOnlineDefrag = 35;

  /// <summary>
  /// This parameter can be used to control the size of the database page cache at run time.
  /// Ordinarily, the cache will automatically tune its size as a function of database and
  /// machine activity levels. If the application sets this parameter to zero, then the cache
  /// will tune its own size in this manner. However, if the application sets this parameter
  /// to a non-zero value then the cache will adjust itself to that target size.
  /// </summary>
  jet_paramCacheSize = 41;

  /// <summary>
  /// When this parameter is true, every database is checked at JetAttachDatabase time for
  /// indexes over Unicode key columns that were built using an older version of the NLS
  /// library in the operating system. This must be done because the database engine persists
  /// the sort keys generated by LCMapStringW and the value of these sort keys change from release to release.
  /// If a primary index is detected to be in this state then JetAttachDatabase will always fail with
  /// JET_err.PrimaryIndexCorrupted.
  /// If any secondary indexes are detected to be in this state then there are two possible outcomes.
  /// If AttachDatabaseGrbit.DeleteCorruptIndexes was passed to JetAttachDatabase then these indexes
  /// will be deleted and JET_wrnCorruptIndexDeleted will be returned from JetAttachDatabase. These
  /// indexes will need to be recreated by your application. If AttachDatabaseGrbit.DeleteCorruptIndexes
  /// was not passed to JetAttachDatabase then the call will fail with JET_errSecondaryIndexCorrupted.
  /// </summary>
  jet_paramEnableIndexChecking = 45;

  /// <summary>
  /// This parameter can be used to control which event log the database engine uses for its event log
  /// messages. By default, all event log messages will go to the Application event log. If the registry
  /// key name for another event log is configured then the event log messages will go there instead.
  /// </summary>
  jet_paramEventSourceKey = 49;

  /// <summary>
  /// When this parameter is true, informational event log messages that would ordinarily be generated by
  /// the database engine will be suppressed.
  /// </summary>
  jet_paramNoInformationEvent = 50;

  /// <summary>
  /// Configures the detail level of eventlog messages that are emitted
  /// to the eventlog by the database engine. Higher numbers will result
  /// in more detailed eventlog messages.
  /// </summary>
  jet_paramEventLoggingLevel = 51;

  /// <summary>
  /// Delete the log files that are not matching (generation wise) during soft recovery.
  /// </summary>
  DeleteOutOfRangeLogs = 52;

  /// <summary>
  /// <para>
  /// After Windows 7, it was discovered that JET_paramEnableIndexCleanup had some implementation limitations, reducing its effectiveness.
  /// Rather than update it to work with locale names, the functionality is removed altogether.
  /// </para>
  /// <para>
  /// Unfortunately JET_paramEnableIndexCleanup can not be ignored altogether. JET_paramEnableIndexChecking defaults to false, so if
  /// JET_paramEnableIndexCleanup were to be removed entirely, then by default there were would be no checks for NLS changes!
  /// </para>
  /// <para>
  /// The current behavious (when enabled) is to track the language sort versions for the indices, and when the sort version for that
  /// particular locale changes, the engine knows which indices are now invalid. For example, if the sort version for only "de-de" changes,
  /// then the "de-de" indices are invalid, but the "en-us" indices will be fine.
  /// </para>
  /// <para>
  /// Post-Windows 8:
  /// JET_paramEnableIndexChecking accepts JET_INDEXCHECKING (which is an enum). The values of '0' and '1' have the same meaning as before,
  /// but '2' is JET_IndexCheckingDeferToOpenTable, which means that the NLS up-to-date-ness is NOT checked when the database is attached.
  /// It is deferred to JetOpenTable(), which may now fail with JET_errPrimaryIndexCorrupted or JET_errSecondaryIndexCorrupted (which
  /// are NOT actual corruptions, but instead reflect an NLS sort change).
  /// </para>
  /// <para>
  /// IN SUMMARY:
  /// New code should explicitly set both IndexChecking and IndexCleanup to the same value.
  /// </para>
  /// </summary>
  jet_paramEnableIndexCleanup = 54;

  /// <summary>
  /// This parameter configures the minimum size of the database page cache. The size is in database pages.
  /// </summary>
  jet_paramCacheSizeMin = 60;

  /// <summary>
  /// This parameter represents a threshold relative to JET_paramMaxVerPages that controls
  /// the discretionary use of version pages by the database engine. If the size of the version store exceeds
  /// this threshold then any information that is only used for optional background tasks, such as reclaiming
  /// deleted space in the database, is instead sacrificed to preserve room for transactional information.
  /// </summary>
  jet_paramPreferredVerPages = 63;

  /// <summary>
  /// This parameter configures the page size for the database. The page
  /// size is the smallest unit of space allocation possible for a database
  /// file. The database page size is also very important because it sets
  /// the upper limit on the size of an individual record in the database.
  /// </summary>
  /// <remarks>
  /// Only one database page size is supported per process at this time.
  /// This means that if you are in a single process that contains different
  /// applications that use the database engine then they must all agree on
  /// a database page size.
  /// </remarks>
  jet_paramDatabasePageSize = 64;

  /// <summary>
  /// This parameter can be used to convert a JET_ERR into a string.
  /// This should only be used with JetGetSystemParameter.
  /// </summary>
  jet_paramErrorToString = 70;

  /// <summary>
  /// Configures the engine with a "JET_CALLBACK" delegate.
  /// This callback may be called for the following reasons:
  /// for more information. This parameter cannot currently be retrieved.
  /// </summary>
  jet_paramRuntimeCallback = 73;

  /// <summary>
  /// This parameter controls the outcome of JetInit when the database
  /// engine is configured to start using transaction log files on disk
  /// that are of a different size than what is configured. Normally,
  /// "Api.JetInit" will successfully recover the databases
  /// but will fail with "JET_err.LogFileSizeMismatchDatabasesConsistent"
  /// to indicate that the log file size is misconfigured. However, when
  /// this parameter is set to true then the database engine will silently
  /// delete all the old log files, start a new set of transaction log files
  /// using the configured log file size. This parameter is useful when the
  /// application wishes to transparently change its transaction log file
  /// size yet still work transparently in upgrade and restore scenarios.
  /// </summary>
  jet_paramCleanupMismatchedLogFiles = 77;

  /// <summary>
  /// This parameter controls what happens when an exception is thrown by the
  /// database engine or code that is called by the database engine. When set
  /// to JET_ExceptionMsgBox, any exception will be thrown to the Windows unhandled
  /// exception filter. This will result in the exception being handled as an
  /// application failure. The intent is to prevent application code from erroneously
  /// trying to catch and ignore an exception generated by the database engine.
  /// This cannot be allowed because database corruption could occur. If the application
  /// wishes to properly handle these exceptions then the protection can be disabled
  /// by setting this parameter to JET_ExceptionNone.
  /// </summary>
  jet_paramExceptionAction = 98;

  /// <summary>
  /// When this parameter is set to true then any folder that is missing in a file system path in use by
  /// the database engine will be silently created. Otherwise, the operation that uses the missing file system
  /// path will fail with JET_err.InvalidPath.
  /// </summary>
  jet_paramCreatePathIfNotExist = 100;

  /// <summary>
  /// When this parameter is true then only one database is allowed to
  /// be opened using JetOpenDatabase by a given session at one time.
  /// The temporary database is excluded from this restriction.
  /// </summary>
  jet_paramOneDatabasePerSession = 102;

  /// <summary>
  /// This parameter controls the maximum number of instances that can be created in a single process.
  /// </summary>
  jet_paramMaxInstances = 104;

  /// <summary>
  /// This parameter controls the number of background cleanup work items that
  /// can be queued to the database engine thread pool at any one time.
  /// </summary>
  jet_paramVersionStoreTaskQueueMax = 105;

  /// <summary>
  /// This parameter controls whether perfmon counters should be enabled or not.
  /// By default, perfmon counters are enabled, but there is memory overhead for enabling
  /// them.
  /// </summary>
  jet_paramDisablePerfmon = 107;

  JET_coltypNil          = 0;  // An invalid column type.
  JET_coltypBit          = 1;  // A column type that allows three values: True, False, or NULL. This type of column is one byte in length and is a fixed size. False sorts before True.
  JET_coltypUnsignedByte = 2;  // A 1-byte unsigned integer that can take on values between 0 (zero) and 255.
  JET_coltypShort        = 3;  // A 2-byte signed integer that can take on values between -32768 and 32767. Negative values sort before positive values.
  JET_coltypLong         = 4;  // A 4-byte signed integer that can take on values between - 2147483648 and 2147483647. Negative values sort before positive values.
  JET_coltypCurrency     = 5;  // An 8-byte signed integer that can take on values between - 9223372036854775808 and 9223372036854775807. Negative values sort before positive values. This column type is identical to the variant currency type. This column type can also be used as a native 8-byte signed integer.
  JET_coltypIEEESingle   = 6;  // A single-precision (4-byte) floating point number.
  JET_coltypIEEEDouble   = 7;  // A double-precision (8-byte) floating point number.
  JET_coltypDateTime     = 8;  // A double-precision (8-byte) floating point number that represents a date in fractional days since the year 1900. This column type is identical to the variant date type.
  JET_coltypBinary       = 9;  // A fixed or variable length, raw binary column that can be up to 255 bytes in length.
  JET_coltypText         = 10; // A fixed or variable length text column that can be up to 255 ASCII characters in length or 127 Unicode characters in length.
  JET_coltypLongBinary   = 11; // A fixed or variable length, raw binary column that can be up to 2147483647 bytes in length. This type is considered to be a Long Value. A Long Value is special because it can be large and because it can be accessed as a stream. This type is otherwise identical to JET_coltypBinary.
  JET_coltypLongText     = 12; // A fixed or variable length, text column that can be up to 2147483647 ASCII characters in length or 1073741823 Unicode characters in length. This type is considered to be a Long Value. A Long Value is special because it can be large and because it can be accessed as a stream. This type is otherwise identical to JET_coltypText.
  JET_coltypSLV          = 13; // This column type is obsolete.
  JET_coltypUnsignedLong = 14; // A 4-byte unsigned integer that can take on values between 0 (zero) and 4294967295.
  JET_coltypLongLong     = 15; // An 8-byte signed integer that can take on values between - 9223372036854775808 and 9223372036854775807. Negative values sort before positive values.
  JET_coltypGUID         = 16; // A fixed length 16 byte binary column that natively represents the GUID data type. GUID column values sort in the same way that those values would sort as strings when in standard form (i.e. {4999b5c0-7657-42d9-bdc1-4b779784e013}).
  JET_coltypUnsignedShort= 17; // A 2-byte unsigned integer that can take on values between 0 and 65535.
  JET_coltypMax          = 18; // A constant describing the maximum (that is, one beyond the largest valid) column type supported by the engine.

  JET_bitColumnFixed         = $00000001; // Is fixed size The column will always use the same size (within the row) regardless of how much data is stored in the column.
  JET_bitColumnTagged        = $00000002; // Is tagged The column is tagged. A tagged columns does not take up any space in the database if it does not contain data.
  JET_bitColumnNotNULL       = $00000004; //  Not empty The column is not allow to be set to an empty value (NULL).
  JET_bitColumnVersion       = $00000008; // Is version column The column is a version column that specifies the version of the row.
  JET_bitColumnAutoincrement = $00000010; // The column will automatically be incremented. The number is an increasing number, and is guaranteed to be unique within a table. The numbers, however, might not be continuous. For example, if five rows are inserted into a table, the "autoincrement" column could contain the values { 1, 2, 6, 7, 8 }. This bit can only be used on columns of type JET_coltypLong or JET_coltypCurrency.
  JET_bitColumnUpdatable     = $00000020; // This bit is valid only on calls to JetGetColumnInfo.
  JET_bitColumnTTKey         = $00000040; // This bit is valid only on calls to JetOpenTable.
  JET_bitColumnTTDescending  = $00000080; // This bit is valid only on calls to JetOpenTempTable.
  JET_bitColumnMultiValued   = $00000400; // The column can be multi-valued. A multi-valued column can have zero, one, or more values associated with it. The various values in a multi-valued column are identified by a number called the itagSequence member, which belongs to various structures, including: JET_RETINFO, JET_SETINFO, JET_SETCOLUMN, JET_RETRIEVECOLUMN, and JET_ENUMCOLUMNVALUE. Multi-valued columns must be tagged columns; that is, they cannot be fixed-length or variable-length columns.
  JET_bitColumnEscrowUpdate  = $00000800; // Specifies that a column is an escrow update column. An escrow update column can be updated concurrently by different sessions with JetEscrowUpdate and will maintain transactional consistency. An escrow update column must also meet the following conditions:
                                          // An escrow update column can be created only when the table is empty.
                                          // An escrow update column must be of type JET_coltypLong.
                                          // An escrow update column must have a default value (that is cbDefault must be positive). JET_bitColumnEscrowUpdate cannot be used in conjunction with JET_bitColumnTagged, JET_bitColumnVersion, or JET_bitColumnAutoincrement.
  JET_bitColumnUnversioned   = $00001000; // The column will be created in an without version information. This means that other transactions that attempt to add a column with the same name will fail. This bit is only useful with JetAddColumn. It cannot be used within a transaction.
  JET_bitColumnDeleteOnZero  = $00002000; // The column is an escrow update column, and when it reaches zero, the record will be deleted. A common use for a column that can be finalized is to use it as a reference count field, and when the field reaches zero the record gets deleted. JET_bitColumnDeleteOnZero is related to JET_bitColumnFinalize. A Delete-on-zero column must be an escrow update column. JET_bitColumnDeleteOnZero cannot be used with JET_bitColumnFinalize. JET_bitColumnDeleteOnZero cannot be used with user defined default columns.
  JET_bitColumnMaybeNull     = $00002000; // Reserved for future use.
  JET_bitColumnFinalize      = $00004000; // Use JET_bitColumnDeleteOnZero instead of JET_bitColumnFinalize. JET_bitColumnFinalize that a column can be finalized. When a column that can be finalized has an escrow update column that reaches zero, the row will be deleted. Future versions might invoke a callback function instead (For more information, see JET_CALLBACK). A column that can be finalized must be an escrow update column. JET_bitColumnFinalize cannot be used with JET_bitColumnUserDefinedDefault.
  JET_bitColumnUserDefinedDefault = $00008000; // The default value for a column will be provided by a callback function. See JET_CALLBACK. A column that has a user-defined default must be a tagged column. Specifying JET_bitColumnUserDefinedDefault means that pvDefault must point to a JET_USERDEFINEDDEFAULT structure, and cbDefault must be set to sizeof( JET_USERDEFINEDDEFAULT ).
                                               // JET_bitColumnUserDefinedDefault cannot be used in conjunction with JET_bitColumnFixed, JET_bitColumnNotNULL, JET_bitColumnVersion, JET_bitColumnAutoincrement, JET_bitColumnUpdatable, JET_bitColumnEscrowUpdate, JET_bitColumnFinalize, JET_bitColumnDeleteOnZero, or JET_bitColumnMaybeNull.

  JET_bitTableDenyRead       = $00000001;  //The table cannot be opened for read-access by another database session.
  JET_bitTableDenyWrite      = $00000002;  //The table cannot be opened for write-access by another database session.
  JET_bitTableNoCache        = $00000004;  //Do not cache the pages for this table.
  JET_bitTablePermitDDL      = $00000008;  //Allows DDL modification on tables flagged as FixedDDL. This option must be used with the JET_bitTableDenyRead option.
  JET_bitTablePreread        = $00000010;  //Provides a hint that the table is probably not in the buffer cache, and that pre-reading may be beneficial to performance.
  JET_bitTableReadOnly       = $00000020;  //Requests read-only access to the table.
  JET_bitTableSequential     = $00000040;  //The table should be very aggressively prefetched from disk because the application will be scanning it sequentially.
  JET_bitTableUpdatable      = $00000080;  //Requests write-access to the table.

  JET_bitIndexUnique         = $00000001; //Duplicate index entries (keys) are not allowed. This is enforced when JetUpdate is called, not when JetSetColumn is called.
  JET_bitIndexPrimary        = $00000002; //The index is a primary (clustered) index. Every table must have exactly one primary index. If no primary index is explicitly defined over a table, the database engine will create its own primary index.
  JET_bitIndexDisallowNull   = $00000004; //None of the columns over which the index is created can contain a NULL value.
  JET_bitIndexIgnoreNull     = $00000008; //Do not add an index entry for a row if all of the columns being indexed are NULL.
  JET_bitIndexIgnoreAnyNull  = $00000020; //Do not add an index entry for a row if any of the columns being indexed are NULL.
  JET_bitIndexIgnoreFirstNull= $00000040; //Do not add an index entry for a row if the first column being indexed is NULL.
  JET_bitIndexLazyFlush      = $00000080; //Index operations will be logged lazily.
  JET_bitIndexEmpty          = $00000100; //Do not attempt to build the index, because all entries would evaluate to NULL. grbit must also specify JET_bitIgnoreAnyNull when JET_bitIndexEmpty is passed. This is a performance enhancement. For example, if a new column is added to a table, an index is created over this newly added column, and all the records in the table are scanned even though they are not added to the index. Specifying JET_bitIndexEmpty skips the scanning of the table, which could potentially take a long time.
  JET_bitIndexUnversioned    = $00000200; //Causes index creation to be visible to other transactions. Typically a session in a transaction will not be able to see an index creation operation in another session. This flag can be useful if another transaction is likely to create the same index. The second index-create will fail instead of potentially causing many unnecessary database operations. The second transaction might not be able to use the index immediately. The index creation operation has to complete before it is usable. The session must not currently be in a transaction to create an index without version information.
  JET_bitIndexSortNullsHigh  = $00000400; //Specifying this flag causes NULL values to be sorted after data for all columns in the index.
  JET_bitIndexUnicode        = $00000800; //Specifying this flag affects the interpretation of the lcid/pidxunicde union field in the structure. Setting the bit means that the pidxunicode field actually points to a JET_UNICODEINDEX structure. JET_bitIndexUnicode is not required in order to index Unicode data. It is only used to customize the normalization of Unicode data.
  JET_bitIndexTuples         = $00001000; //Specifies that the index is a tuple index. See JET_TUPLELIMITS for a description of a tuple index.
  JET_bitIndexTupleLimits    = $00002000; //Specifying this flag affects the interpretation of the cbVarSegMac/ptuplelimits union field in the structure. Setting this bit means that the ptuplelimits field actually points to a JET_TUPLELIMITS structure to allow custom tuple index limits (implies JET_bitIndexTuples).
  JET_bitIndexCrossProduct   = $00004000; //Specifying this flag for an index that has more than one key column that is a multivalued column will result in an index entry being created for each result of a cross product of all the values in those key columns. Otherwise, the index would only have one entry for each multivalue in the most significant key column that is a multivalued column and each of those index entries would use the first multivalue from any other key columns that are multivalued columns.
  JET_bitIndexKeyMost        = $00008000; //Specifying this flag will cause the index to use the maximum key size specified in the cbKeyMost field in the structure. Otherwise, the index will use JET_cbKeyMost (255) as its maximum key size.
  JET_bitIndexDisallowTruncation= $00010000; //Specifying this flag will cause any update to the index that would result in a truncated key to fail with JET_errKeyTruncated. Otherwise, keys will be silently truncated. For more information on key truncation, see the JetMakeKey function.
  JET_bitIndexNestedTable    = $00020000; //Specifying this flag will cause update the index over multiple multivalued columns but only with values of same itagSequence.

  /// <summary>
  /// Perform recovery, but halt at the Undo phase. Allows whatever logs are present to
  /// be replayed, then later additional logs can be copied and replayed.
  /// </summary>
  JET_bitRecoveryWithoutUndo = $08;

  /// <summary>
  /// On successful soft recovery, truncate log files.
  /// </summary>
  JET_bitTruncateLogsAfterRecovery = $10;

  /// <summary>
  /// Missing database map entry default to same location.
  /// </summary>
  JET_bitReplayMissingMapEntryDB = $20;

  /// <summary>
  /// Transaction logs must exist in the log file directory
  /// (i.e. can't auto-start a new stream).
  /// </summary>
  JET_bitLogStreamMustExist = $40;

  /// <summary>
  /// Recover without error even if uncommitted logs have been lost. Set
  /// the recovery waypoint with Windows7Param.WaypointLatency to enable
  /// this type of recovery.
  /// </summary>
  JET_bitReplayIgnoreLostLogs = $80;

  JET_MoveFirst              = -Int64(2147483648);
  JET_MovePrevious           = -1;
  JET_MoveNext               = 1;
  JET_MoveLast               = $7fffffff;

  JET_prepInsert             = 0;
  JET_prepReplace            = 2;
  JET_prepCancel             = 3;
  JET_prepReplaceNoLock      = 4;
  JET_prepInsertCopy         = 5;
  JET_prepInsertCopyDeleteOriginal = 7;

  JET_bitSetAppendLV         = $1;
  JET_bitSetOverwriteLV      = $4;
  JET_bitSetRevertToDefaultValue = $200;
  JET_bitSetSeparateLV       = $40;
  JET_bitSetSizeLV           = $8;
  JET_bitSetUniqueMultiValues = $80;
  JET_bitSetUniqueNormalizedMultiValues = $100;
  JET_bitSetZeroLength       = $20;
  JET_bitSetIntrinsicLV      = $400;

  JET_objtypNil              = 0;
  JET_objtypTable            = 1;

  JET_bitNone                = 0;
  JET_bitNewKey              = $1;
  JET_bitStrLimit            = $2;
  JET_bitSubStrLimit         = $4;
  JET_bitNormalizedKey       = $8;
  JET_bitKeyDataZeroLength   = $10;
  JET_bitFullColumnStartLimit= $100;
  JET_bitFullColumnEndLimit  = $200;
  JET_bitPartialColumnStartLimit = $400;
  JET_bitPartialColumnEndLimit = $800;

  JET_bitSeekEQ              = $1;
  JET_bitSeekLT              = $2;
  JET_bitSeekLE              = $4;
  JET_bitSeekGE              = $8;
  JET_bitSeekGT              = $10;
  JET_bitSetIndexRange       = $20;

  JET_bitRangeInclusive      = $1;
  JET_bitRangeUpperLimit     = $2;
  JET_bitRangeInstantDuration= $4;
  JET_bitRangeRemove         = $8;

  JET_bitDefragmentAvailSpaceTreesOnly = $40;
  JET_bitDefragmentBatchStart = $1;
  JET_bitDefragmentBatchStop = $2;
  JET_bitDefragmentNoPartialMerges = $80;
  JET_bitDefragmentBTree = $100;

  JET_wrnNyi = -1;
  JET_errRfsFailure = -100;
  JET_errRfsNotArmed = -101;
  JET_errFileClose = -102;
  JET_errOutOfThreads = -103;
  JET_errTooManyIO = -105;
  JET_errTaskDropped = -106;
  JET_errInternalError = -107;
  JET_errDatabaseBufferDependenciesCorrupted = -255;
  JET_errPreviousVersion = -322;
  JET_errPageBoundary = -323;
  JET_errKeyBoundary = -324;
  JET_errBadPageLink = -327;
  JET_errBadBookmark = -328;
  JET_errNTSystemCallFailed = -334;
  JET_errBadParentPageLink = -338;
  JET_errSPAvailExtCacheOutOfSync = -340;
  JET_errSPAvailExtCorrupted = -341;
  JET_errSPAvailExtCacheOutOfMemory = -342;
  JET_errSPOwnExtCorrupted = -343;
  JET_errDbTimeCorrupted = -344; //
  JET_errKeyTruncated = -346; //
  JET_errKeyTooBig = -408; //
  JET_errInvalidLoggedOperation = -500;  //The logged operation cannot be redone
  JET_errLogFileCorrupt = -501;  //The log file is corrupt
  JET_errNoBackupDirectory = -503;  //A backup directory was not given
  JET_errBackupDirectoryNotEmpty = -504;  //The backup directory is not empty
  JET_errBackupInProgress = -505;  //The backup is active already
  JET_errRestoreInProgress = -506; //A restore is in progress
  JET_errMissingPreviousLogFile = -509;  //The log file is missing for the check point
  JET_errLogWriteFail = -510; //There was a failure writing to the log file
  JET_errLogDisabledDueToRecoveryFailure = -511; //The attempt to write to the log after recovery failed
  JET_errCannotLogDuringRecoveryRedo = -512; //The attempt to write to the log during the recovery redo failed
  JET_errLogGenerationMismatch = -513; //The name of of the log file does not match the internal generation number
  JET_errBadLogVersion = -514; //The version of the log file is not compatible with the ESE version
  JET_errInvalidLogSequence = -515; //The timestamp in the next log does not match the expected timestamp
  JET_errLoggingDisabled = -516; //The log is not active
  JET_errLogBufferTooSmall = -517; //The log buffer is too small for recovery
  JET_errLogSequenceEnd = -519; //The maximum log file number has been exceeded
  JET_errNoBackup = -520; //There is no backup in progress
  JET_errInvalidBackupSequence = -521; //The backup call is out of sequence
  JET_errBackupNotAllowedYet = -523; //A backup cannot be done at this time
  JET_errDeleteBackupFileFail = -524; //A backup file could not be deleted
  JET_errMakeBackupDirectoryFail = -525; //The backup temporary directory could not be created
  JET_errInvalidBackup = -526; //Circular logging is enabled; an incremental backup cannot be performed
  JET_errRecoveredWithErrors = -527; //The data was restored with errors
  JET_errMissingLogFile = -528; //The current log file is missing
  JET_errLogDiskFull = -529; //The log disk is full
  JET_errBadLogSignature = -530; //There is a bad signature for a log file
  JET_errBadDbSignature = -531; //There is a bad signature for a database file
  JET_errBadCheckpointSignature = -532; //There is a bad signature for a checkpoint file
  JET_errCheckpointCorrupt = -533; //The checkpoint file was not found or was corrupt
  JET_errMissingPatchPage = -534; //The database patch file page was not found during recovery
  JET_errBadPatchPage = -535; //The database patch file page is not valid
  JET_errRedoAbruptEnded = -536; //The redo abruptly ended due to a sudden failure while reading logs from the log file
  JET_errBadSLVSignature = -537; //This flag is reserved
  JET_errPatchFileMissing = -538; //The hard restore detected that a database patch file is missing from the backup set
  JET_errDatabaseLogSetMismatch = -539; //The database does not belong with the current set of log files
  JET_errDatabaseStreamingFileMismatch = -540; //This flag is reserved
  JET_errLogFileSizeMismatch = -541; //The actual log file size does not match JET_paramLogFileSize
  JET_errCheckpointFileNotFound = -542; //The checkpoint file could not be located
  JET_errRequiredLogFilesMissing = -543; //The required log files for recovery are missing
  JET_errSoftRecoveryOnBackupDatabase = -544; //A soft recovery is about to be used on a backup database when a restore should be used instead
  JET_errLogFileSizeMismatchDatabasesConsistent = -545; //The databases have been recovered, but the log file size used during recovery does not match JET_paramLogFileSize
  JET_errLogSectorSizeMismatch = -546; //The log file sector size does not match the sector size of the current volume
  JET_errLogSectorSizeMismatchDatabasesConsistent = -547; //The databases have been recovered, but the log file sector size (used during recovery) does not match the sector size of the current volume
  JET_errLogSequenceEndDatabasesConsistent = -548; //The databases have been recovered, but all possible log generations in the current sequence have been used. All log files and the checkpoint file must be deleted and databases must be backed up before continuing
  JET_errStreamingDataNotLogged = -549; //There was an illegal attempt to replay a streaming file operation where the data was not logged. This is probably caused by an attempt to rollforward with circular logging enabled
  JET_errDatabaseDirtyShutdown = -550; //The database was not shutdown cleanly. A recovery must first be run to properly complete database operations for the previous shutdown.
  JET_errConsistentTimeMismatch = -551; //The last consistent time for the database has not been matched.
  JET_errDatabasePatchFileMismatch = -552; //The database patch file is not generated from this backup.
  JET_errEndingRestoreLogTooLow = -553; //The starting log number is too low for the restore.
  JET_errStartingRestoreLogTooHigh = -554; //The starting log number is too high for the restore.
  JET_errGivenLogFileHasBadSignature = -555; //The restore log file has a bad signature.
  JET_errGivenLogFileIsNotContiguous = -556; //The restore log file is not contiguous.
  JET_errMissingRestoreLogFiles = -557; //Some restore log files are missing.
  JET_errMissingFullBackup = -560; //The database missed a previous full backup before attempting to perform an incremental backup.
  JET_errBadBackupDatabaseSize = -561; //The backup database size is not a multiple of the database page size.
  JET_errDatabaseAlreadyUpgraded = -562; //The current attempt to upgrade a database has been stopped because the database is already current.
  JET_errDatabaseIncompleteUpgrade = -563; //The database was only partially converted to the current format. The database must be restored from backup.
  JET_errMissingCurrentLogFiles = -565; //Some current log files are missing for continuous restore.
  JET_errDbTimeTooOld = -566; //The dbtime on a page is smaller than the dbtimeBefore that is in the record.
  JET_errDbTimeTooNew = -567; //The dbtime on a page is in advance of the dbtimeBefore that is in the record.
  JET_errMissingFileToBackup = -569; //Some log or database patch files were missing during the backup.
  JET_errLogTornWriteDuringHardRestore = -570; //A torn write was detected in a backup that was set during a hard restore.
  JET_errLogTornWriteDuringHardRecovery = -571; //A torn write was detected during a hard recovery (the log was not part of a backup set).
  JET_errLogCorruptDuringHardRestore = -573; //Corruption was detected in a backup set during a hard restore.
  JET_errLogCorruptDuringHardRecovery = -574; //Corruption was detected during hard recovery (the log was not part of a backup set).
  JET_errMustDisableLoggingForDbUpgrade = -575; //Logging cannot be enabled while attempting to upgrade a database.
  JET_errBadRestoreTargetInstance = -577; //Either the TargetInstance that was specified for restore has not been found or the log files do not match.
  JET_errRecoveredWithoutUndo = -579; //The database engine successfully replayed all operations in the transaction log to perform a crash recovery but the caller elected to stop recovery without rolling back uncommitted updates.
  JET_errDatabasesNotFromSameSnapshot = -580; //The databases to be restored are not from the same shadow copy backup.
  JET_errSoftRecoveryOnSnapshot = -581; //There is a soft recovery on a database from a shadow copy backup set.
  JET_errUnicodeTranslationBufferTooSmall = -601; //The Unicode translation buffer is too small.
  JET_errUnicodeTranslationFail = -602; //The Unicode normalization failed.
  JET_errUnicodeNormalizationNotSupported = -603; //The operating system does not provide support for Unicode normalization and a normalization callback was not specified.
  JET_errExistingLogFileHasBadSignature = -610; //The existing log file has a bad signature.
  JET_errExistingLogFileIsNotContiguous = -611; //An existing log file is not contiguous.
  JET_errLogReadVerifyFailure = -612; //A checksum error was found in the log file during backup.
  JET_errSLVReadVerifyFailure = -613; //This flag is reserved.
  JET_errCheckpointDepthTooDeep = -614; //There are too many outstanding generations between the checkpoint and the current generation.
  JET_errRestoreOfNonBackupDatabase = -615; //A hard recovery was attempted on a database that was not a backup database.
  JET_errInvalidGrbit = -900; //There is an invalid grbit parameter.
  JET_errTermInProgress = -1000; //Termination is in progress.
  JET_errFeatureNotAvailable = -1001; //This API element is not supported.
  JET_errInvalidName = -1002; //An invalid name is being used.
  JET_errInvalidParameter = -1003; //An invalid API parameter is being used.
  JET_errDatabaseFileReadOnly = -1008; //There was an attempt to attach to a read-only database file for read/write operations.
  JET_errInvalidDatabaseId = -1010; //There is an invalid database ID.
  JET_errOutOfMemory = -1011; //The system is out of memory.
  JET_errOutOfDatabaseSpace = -1012; //The maximum database size has been reached.
  JET_errOutOfCursors = -1013; //The table is out of cursors.
  JET_errOutOfBuffers = -1014; //The database is out of page buffers.
  JET_errTooManyIndexes = -1015; //There are too many indexes.
  JET_errTooManyKeys = -1016; //There are too many columns in an index.
  JET_errRecordDeleted = -1017; //The record has been deleted.
  JET_errReadVerifyFailure = -1018; //There is a checksum error on a database page.
  JET_errPageNotInitialized = -1019; //There is a blank database page.
  JET_errOutOfFileHandles = -1020; //There are no file handles.
  JET_errDiskIO = -1022; //There is a disk IO error.
  JET_errInvalidPath = -1023; //There is an invalid file path.
  JET_errInvalidSystemPath = -1024; //There is an invalid system path.
  JET_errInvalidLogDirectory = -1025; //There is an invalid log directory.
  JET_errRecordTooBig = -1026; //The record is larger than maximum size.
  JET_errTooManyOpenDatabases = -1027; //There are too many open databases.
  JET_errInvalidDatabase = -1028; //This is not a database file.
  JET_errNotInitialized = -1029; //The database engine has not been initialized.
  JET_errAlreadyInitialized = -1030; //The database engine is already initialized.
  JET_errInitInProgress = -1031; //The database engine is being initialized.
  JET_errFileAccessDenied = -1032; //The file cannot be accessed because the file is locked or in use.
  JET_errBufferTooSmall = -1038; //The buffer is too small.
  JET_errTooManyColumns = -1040; //Too many columns are defined.
  JET_errContainerNotEmpty = -1043; //The container is not empty.
  JET_errInvalidFilename = -1044; //The filename is invalid.
  JET_errInvalidBookmark = -1045; //There is an invalid bookmark.
  JET_errColumnInUse = -1046; //The column used is in an index.
  JET_errInvalidBufferSize = -1047; //The data buffer does not match the column size.
  JET_errColumnNotUpdatable = -1048; //The column value cannot be set.
  JET_errIndexInUse = -1051; //The index is in use.
  JET_errLinkNotSupported = -1052; //The link support is unavailable.
  JET_errNullKeyDisallowed = -1053; //Null keys are not allowed on an index.
  JET_errNotInTransaction = -1054; //The operation must occur within a transaction.
  JET_errTooManyActiveUsers = -1059; //There are too many active database users
  JET_errInvalidCountry = -1061; //There is an invalid or unknown country code.
  JET_errInvalidLanguageId = -1062; //There is an invalid or unknown language ID.
  JET_errInvalidCodePage = -1063; //There is an invalid or unknown code page.
  JET_errInvalidLCMapStringFlags = -1064; //There are invalid flags being used for LCMapString.
  JET_errVersionStoreEntryTooBig = -1065; //There was an attempt to create a version store entry (RCE) that was larger than a version bucket.
  JET_errVersionStoreOutOfMemoryAndCleanupTimedOut = -1066; //The version store is out of memory and the cleanup attempt failed to complete.
  JET_errVersionStoreOutOfMemory = -1069; //The version store is out of memory and a cleanup was already attempted).
  JET_errCannotIndex = -1071; //The escrow and SLV columns cannot be indexed.
  JET_errRecordNotDeleted = -1072; //The record has not been deleted.
  JET_errTooManyMempoolEntries = -1073; //Too many mempool entries have been requested.
  JET_errOutOfObjectIDs = -1074; //The database is out of B+ tree ObjectIDs so an offline defragmentation must be performed to reclaim freed or unused ObjectIds.
  JET_errOutOfLongValueIDs = -1075; //The Long-value ID counter has reached the maximum value. An offline defragmentation must be performed to reclaim free or unused LongValueIDs.
  JET_errOutOfAutoincrementValues = -1076; //The auto-increment counter has reached the maximum value. An offline defragmentation will not be able to reclaim free or unused auto-increment values).
  JET_errOutOfDbtimeValues = -1077; //The Dbtime counter has reached the maximum value. An offline defragmentation must be performed to reclaim free or unused Dbtime values.
  JET_errOutOfSequentialIndexValues = -1078; //A sequential index counter has reached the maximum value. An offline defragmentation must be performed to reclaim free or unused SequentialIndex values.
  JET_errRunningInOneInstanceMode = -1080; //This multi-instance call has the single-instance mode enabled.
  JET_errRunningInMultiInstanceMode =-1081; //This single-instance call has the multi-instance mode enabled.
  JET_errSystemParamsAlreadySet = -1082; //The global system parameters have already been set.
  JET_errSystemPathInUse = -1083; //The system path is already being used by another database instance.
  JET_errLogFilePathInUse = -1084; //The log file path is already being used by another database instance.
  JET_errTempPathInUse = -1085; //The path to the temporary database is already being used by another database instance.
  JET_errInstanceNameInUse = -1086; //The instance name is already in use.
  JET_errInstanceUnavailable = -1090; //This instance cannot be used because it encountered a fatal error.
  JET_errDatabaseUnavailable = -1091; //This database cannot be used because it encountered a fatal error.
  JET_errInstanceUnavailableDueToFatalLogDiskFull = -1092; //This instance cannot be used because it encountered a log-disk-full error while performing an operation (such as a transaction rollback) that could not tolerate failure.
  JET_errOutOfSessions = -1101; //The database is out of sessions.
  JET_errWriteConflict = -1102; //The write lock failed due to the existence of an outstanding write lock.
  JET_errTransTooDeep = -1103; //The transactions are nested too deeply.
  JET_errInvalidSesid = -1104; //There is an invalid session handle.
  JET_errWriteConflictPrimaryIndex = -1105; //An update was attempted on an uncommitted primary index.
  JET_errInTransaction = -1108; //The operation is not allowed within a transaction.
  JET_errRollbackRequired = -1109; //The current transaction must be rolled back. It cannot be committed and a new one cannot be started.
  JET_errTransReadOnly = -1110; //A read-only transaction tried to modify the database.
  JET_errSessionWriteConflict = -1111; //There was an attempt to replace the same record by two different cursors in the same session.
  JET_errRecordTooBigForBackwardCompatibility = -1112; //The record would be too big if represented in a database format from a previous version of Jet.
  JET_errCannotMaterializeForwardOnlySort = -1113; //The temporary table could not be created due to parameters that conflict with JET_bitTTForwardOnly.
  JET_errSesidTableIdMismatch = -1114; //The session handle cannot be used with the table id because it was not used to create it.
  JET_errInvalidInstance = -1115; //The instance handle is invalid or refers to an instance that has been shut down.
  JET_errReadLostFlushVerifyFailure = -1119; //The database page read from disk had a previous write not represented on the page. Available on Windows 8 and later for client, and Windows Server 2012 and later for server.
  JET_errDatabaseDuplicate = -1201; //The database already exists.
  JET_errDatabaseInUse = -1202; //The database in use.
  JET_errDatabaseNotFound = -1203; //There is no such database.
  JET_errDatabaseInvalidName = -1204; //The database name is invalid.
  JET_errDatabaseInvalidPages = -1205; //There are an invalid number of pages.
  JET_errDatabaseCorrupted = -1206; //There is a non-database file or corrupt database.
  JET_errDatabaseLocked = -1207; //The database is exclusively locked.
  JET_errCannotDisableVersioning = -1208; //The versioning for this database cannot be disabled.
  JET_errInvalidDatabaseVersion = -1209; //The database engine is incompatible with the database.
  JET_errDatabase200Format = -1210; //The database is in an older (200) format. This error is returned by JetInit if JET_paramCheckFormatWhenOpenFail is set. Windows NT client only.
  JET_errDatabase400Format = -1211; //The database is in an older (400) format. This error is returned by JetInit if JET_paramCheckFormatWhenOpenFail is set. Windows NT client only.
  JET_errDatabase500Format = -1212; //The database is in an older (500) format. This error is returned by JetInit if JET_paramCheckFormatWhenOpenFail is set. Windows NT client only.
  JET_errPageSizeMismatch = -1213; //The database page size does not match the engine.
  JET_errTooManyInstances = -1214; //No more database instances can be started.
  JET_errDatabaseSharingViolation = -1215; //A different database instance is using this database.
  JET_errAttachedDatabaseMismatch = -1216; //An outstanding database attachment has been detected at the start or end of the recovery, but the database is missing or does not match attachment info.
  JET_errDatabaseInvalidPath = -1217; //The specified path to the database file is illegal.
  JET_errDatabaseIdInUse = -1218; //A database is being assigned an ID that is already in use.
  JET_errForceDetachNotAllowed = -1219; //The force detach is allowed only after the normal detach was stopped due to an error.
  JET_errCatalogCorrupted = -1220; //Corruption was detected in the catalog.
  JET_errPartiallyAttachedDB = -1221; //The database is only partially attached and the attach operation cannot be completed.
  JET_errDatabaseSignInUse = -1222; //The database with the same signature is already in use.
  JET_errDatabaseCorruptedNoRepair = -1224; //The database is corrupted but a repair is not allowed.
  JET_errInvalidCreateDbVersion = -1225; //The database engine attempted to replay a Create Database operation from the transaction log but failed due to an incompatible version of that operation.
  JET_errTableLocked = -1302; //The table is exclusively locked.
  JET_errTableDuplicate = -1303; //The table already exists.
  JET_errTableInUse = -1304; //The table is in use and cannot be locked.
  JET_errObjectNotFound = -1305; //There is no such table or object.
  JET_errDensityInvalid = -1307; //There is a bad file or index density.
  JET_errTableNotEmpty = -1308; //The table is not empty.
  JET_errInvalidTableId = -1310; //The table ID is invalid.
  JET_errTooManyOpenTables = -1311; //No more tables can be opened, even after the internal cleanup task has run.
  JET_errIllegalOperation = -1312; //The operation is not supported on the table.
  JET_errTooManyOpenTablesAndCleanupTimedOut = -1313; //No more tables can be opened because the cleanup attempt failed to complete.
  JET_errObjectDuplicate = -1314; //The table or object name is in use.
  JET_errInvalidObject = -1316; //The object is invalid for operation.
  JET_errCannotDeleteTempTable = -1317; //JetCloseTable must be used instead of JetDeleteTable to delete a temporary table.
  JET_errCannotDeleteSystemTable = -1318; //There was an illegal attempt to delete a system table.
  JET_errCannotDeleteTemplateTable = -1319; //There was an illegal attempt to delete a template table.
  JET_errExclusiveTableLockRequired = -1322; //There must be an exclusive lock on the table.
  JET_errFixedDDL = -1323; //DDL operations are prohibited on this table.
  JET_errFixedInheritedDDL = -1324; //On a derived table, DDL operations are prohibited on the inherited portion of the DDL.
  JET_errCannotNestDDL = -1325; //Nesting the hierarchical DDL is not currently supported.
  JET_errDDLNotInheritable = -1326; //There was an attempt to inherit a DDL from a table that is not marked as a template table.
  JET_errInvalidSettings = -1328; //The system parameters were set improperly.
  JET_errClientRequestToStopJetService = -1329; //The client has requested that the service be stopped.
  JET_errCannotAddFixedVarColumnToDerivedTable = -1330; //The Template table was created with the NoFixedVarColumnsInDerivedTables flag set.
  JET_errIndexCantBuild = -1401; //The index build failed.
  JET_errIndexHasPrimary = -1402; //The primary index is already defined.
  JET_errIndexDuplicate = -1403; //The index is already defined.
  JET_errIndexNotFound = -1404; //There is no such index.
  JET_errIndexMustStay = -1405; //The clustered index cannot be deleted.
  JET_errIndexInvalidDef = -1406; //The index definition is invalid.
  JET_errInvalidCreateIndex = -1409; //The creation of the index description was invalid.
  JET_errTooManyOpenIndexes = -1410; //The database is out of index description blocks.
  JET_errMultiValuedIndexViolation = -1411; //Non-unique inter-record index keys have been generated for a multi-valued index.
  JET_errIndexBuildCorrupted = -1412; //A secondary index that properly reflects the primary index failed to build.
  JET_errPrimaryIndexCorrupted = -1413; //The primary index is corrupt and the database must be defragmented.
  JET_errSecondaryIndexCorrupted = -1414; //The secondary index is corrupt and the database must be defragmented.
  JET_errInvalidIndexId = -1416; //The index ID is invalid.
  JET_errIndexTuplesSecondaryIndexOnly = -1430; //The tuple index can only be set on a secondary index.
  JET_errIndexTuplesTooManyColumns = -1431; //The index definition for the tuple index contains more key columns that the database engine can support.
  JET_errIndexTuplesNonUniqueOnly = -1432; //The tuple index must be a non-unique index.
  JET_errIndexTuplesTextBinaryColumnsOnly = -1433; //A tuple index definition can only contain key columns that have text or binary column types.
  JET_errIndexTuplesVarSegMacNotAllowed = -1434; //The tuple index does not allow setting cbVarSegMac.
  JET_errIndexTuplesInvalidLimits = -1435; //The minimum/maximum tuple length or the maximum number of characters that are specified for an index are invalid.
  JET_errIndexTuplesCannotRetrieveFromIndex = -1436; //JetRetrieveColumn cannot be called with the JET_bitRetrieveFromIndex flag set while retrieving a column on a tuple index.
  JET_errIndexTuplesKeyTooSmall = -1437; //The specified key does not meet the minimum tuple length.
  JET_errColumnLong = -1501; //The column value is long.
  JET_errColumnNoChunk = -1502; //There is no such chunk in a long value.
  JET_errColumnDoesNotFit = -1503; //The field will not fit in the record.
  JET_errNullInvalid = -1504; //Null is not valid.
  JET_errColumnIndexed = -1505; //The column is indexed and cannot be deleted.
  JET_errColumnTooBig = -1506; //The field length is greater than maximum allowed length.
  JET_errColumnNotFound = -1507; //There is no such column.
  JET_errColumnDuplicate = -1508; //This field is already defined.
  JET_errMultiValuedColumnMustBeTagged = -1509; //An attempt was made to create a multi-valued column, but the column was not tagged.
  JET_errColumnRedundant = -1510; //There was a second auto-increment or version column.
  JET_errInvalidColumnType = -1511; //The column data type is invalid.
  JET_errTaggedNotNULL = -1514; //There are no non-NULL tagged columns.
  JET_errNoCurrentIndex = -1515; //The database is invalid because it does not contain a current index.
  JET_errKeyIsMade = -1516; //The key is completely made.
  JET_errBadColumnId = -1517; //The column ID is incorrect.
  JET_errBadItagSequence = -1518; //There is a bad itagSequence for the tagged column.
  JET_errColumnInRelationship = -1519; //A column cannot be deleted because it is part of a relationship.
  JET_errCannotBeTagged = -1521; //The auto increment and version cannot be tagged.
  JET_errDefaultValueTooBig = -1524; //The default value exceeds the maximum size.
  JET_errMultiValuedDuplicate = -1525; //A duplicate value was detected on a unique multi-valued column.
  JET_errLVCorrupted = -1526; //Corruption was encountered in a long-value tree.
  JET_errMultiValuedDuplicateAfterTruncation = -1528; //A duplicate value was detected on a unique multi-valued column after the data was normalized, and it is normalizing truncated the data before comparison.
  JET_errDerivedColumnCorruption = -1529; //There is an invalid column in derived table.
  JET_errInvalidPlaceholderColumn = -1530; //An attempt was made to convert a column to a primary index placeholder, but the column does not meet the necessary criteria.
  JET_errRecordNotFound = -1601; //The key was not found.
  JET_errRecordNoCopy = -1602; //There is no working buffer.
  JET_errNoCurrentRecord = -1603; //There is no current record.
  JET_errRecordPrimaryChanged = -1604; //The primary key might not change.
  JET_errKeyDuplicate = -1605; //There is an illegal duplicate key.
  JET_errAlreadyPrepared = -1607; //An attempt was made to update a record while a record update was already in progress.
  JET_errKeyNotMade = -1608; //A call was not made to JetMakeKey.
  JET_errUpdateNotPrepared = -1609; //A call was not made to JetPrepareUpdate.
  JET_errDataHasChanged = -1611; //The data has changed and the operation was aborted.
  JET_errLanguageNotSupported = -1619; //This Windows installation does not support the selected language.
  JET_errTooManySorts = -1701; //There are too many sort processes.
  JET_errInvalidOnSort = -1702; //An invalid operation occurred during a sort.
  JET_errTempFileOpenError = -1803; //The temporary file could not be opened.
  JET_errTooManyAttachedDatabases = -1805; //Too many databases are open.
  JET_errDiskFull = -1808; //There is no space left on disk.
  JET_errPermissionDenied = -1809; //Permission is denied.
  JET_errFileNotFound = -1811; //The file was not found.
  JET_errFileInvalidType = -1812; //The file type is invalid.
  JET_errAfterInitialization = -1850; //A restore cannot be started after initialization.
  JET_errLogCorrupted = -1852; //The logs could not be interpreted.
  JET_errInvalidOperation = -1906; //The operation is invalid.
  JET_errAccessDenied = -1907; //Access is denied.
  JET_errTooManySplits = -1909; //An infinite split.
  JET_errSessionSharingViolation = -1910; //Multiple threads are using the same session.
  JET_errEntryPointNotFound = -1911; //An entry point in a required DLL could not be found.
  JET_errSessionContextAlreadySet = -1912; //The specified session already has a session context set.
  JET_errSessionContextNotSetByThisThread = -1913; //An attempt was made to reset the session context, but the current thread was not the original one that set the session context.
  JET_errSessionInUse = -1914; //An attempt was made to terminate the session currently in use.
  JET_errRecordFormatConversionFailed = -1915; //An internal error occurred during a dynamic record format conversion.
  JET_errOneDatabasePerSession = -1916; //Only one open user database per session is allowed (as indicated by setting the JET_paramOneDatabasePerSession flag during database creation).
  JET_errRollbackError = -1917; //There was an error during rollback.
  JET_errCallbackFailed = -2101; //A callback function call failed.
  JET_errCallbackNotResolved = -2102; //A callback function could not be found.
  JET_errOSSnapshotInvalidSequence = -2401; //The operating system shadow copy API was used in an invalid sequence.
  JET_errOSSnapshotTimeOut = -2402; //The operating system shadow copy ended with a time-out.
  JET_errOSSnapshotNotAllowed = -2403; //The operating system shadow copy is not allowed because a backup or recovery in is progress.
  JET_errOSSnapshotInvalidSnapId = -2404; //The operation failed because the specified operating system shadow copy handle was invalid.
  JET_errLSCallbackNotSpecified = -3000; //An attempt was made to use local storage without a callback function being specified.
  JET_errLSAlreadySet = -3001; //An attempt was made to set the local storage for an object which already had it set.
  JET_errLSNotSet = -3002; //An attempt was made to retrieve local storage from an object which did not have it set.
  JET_errFileIOSparse = -4000; //An I/O operation failed because it was attempted against an unallocated region of a file.
  JET_errFileIOBeyondEOF = -4001; //A read was issued to a location beyond the EOF (writes will expand the file).
  JET_errFileIOAbort = -4002; //This flag instructs the JET_ABORTRETRYFAILCALLBACK caller to abort the specified I/O.
  JET_errFileIORetry = -4003; //This flag instructs the JET_ABORTRETRYFAILCALLBACK caller to retry the specified I/O.
  JET_errFileIOFail = -4004; //This flag instructs the JET_ABORTRETRYFAILCALLBACK caller to fail the specified I/O.
  JET_errFileCompressed = -4005; //Read/write access is not supported on compressed files.

implementation

end.
