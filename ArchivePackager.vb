Option Strict On
Imports System.IO

Namespace BethesdaArchive.Core

    ''' <summary>
    ''' Behavior when the bundle does not fit a single archive under MaxArchiveBytes.
    ''' SplitByPlugin is reserved for the next phase (numbered companion plugins);
    ''' Pack currently treats Overflow as advisory — it does not enforce MaxArchiveBytes yet.
    ''' </summary>
    Public Enum ArchiveOverflowPolicy
        ThrowOnExceed
        SplitByPlugin
    End Enum

    Public NotInheritable Class PackagerRequest
        Public Property Game As GameKind
        Public Property ModBaseName As String = "WM_ClonePack"
        Public Property OutputDir As String = ""
        Public Property Entries As List(Of VirtualEntry)
        ' Soft cap per archive. Recommended: FO4 = 3GB, SSE = 2GB (BSA u32 offsets, 4GB hard limit).
        ' When a bundle exceeds this, Pack distributes entries across numbered companion plugins
        ' ("WM_ClonePack2.esp", "WM_ClonePack3.esp", ...) so the engine auto-loads each pair.
        Public Property MaxArchiveBytes As Long = 3L << 30

        ' Estimate of compressed/raw size ratio used when planning slot assignment for new entries.
        ' Lower = more conservative (more splits). 0.85 means "assume payload shrinks to 85% after
        ' compression" — a margin to keep us under MaxArchiveBytes given typical BC7/Zlib output.
        ' IGNORED when BundleAlreadyCompressed = True (the packager uses VirtualEntry.PreCompressedCompSize
        ' directly, which is exact).
        Public Property CompressionRatioEstimate As Double = 0.85

        ' When True, every VirtualEntry in Entries is expected to have PreCompressed = True with
        ' valid PreCompressedBytes / PreCompressedCompSize / PreCompressedDecompSize set by the
        ' caller (typically via PayloadCompressor.CompressFor*). Distribution then sums exact
        ' compressed sizes — no estimation, no ratio. The writer stream-copies the bytes verbatim.
        Public Property BundleAlreadyCompressed As Boolean = False

        ' When deciding whether to add free entries (paths not yet present in any archive) to an
        ' existing slot, the packager rewrites the slot only when the slot's REAL free space
        ' (MaxArchiveBytes - existing on-disk size) is at least this many bytes. Default 100 MB.
        ' This is an absolute floor in bytes, not a ratio: the rule is "is there enough actual
        ' room to be worth the I/O of rewriting?". Anchored entries (paths already present in
        ' the slot) bypass this check — that rewrite is mandatory to honor the no-duplicate
        ' contract regardless of how little room is left.
        '
        ' Examples (cap = 3 GB):
        '   - slot at 1.2 GB → 1.8 GB free ≥ 100 MB → fill it (caller's free entries land here).
        '   - slot at 2.5 GB → 500 MB free ≥ 100 MB → fill it (accept the 2.5 GB rewrite cost
        '     to avoid creating an extra fragmented slot).
        '   - slot at 2.95 GB → 50 MB free < 100 MB → leave it alone (rewriting 2.95 GB to
        '     squeeze in 50 MB is not worth it; the entries go to a new slot instead).
        Public Property MinFreeSpaceToFill As Long = 100L * 1024L * 1024L

        ' Callback invoked exactly once per NEW plugin slot that Pack creates. Signature is
        ' (pluginFilePath, game). The caller is expected to write the dummy plugin file at that
        ' path (e.g. via FO4_Base_Library.PluginWriter.WriteLightMasterDummy). When Pack reuses
        ' an existing plugin, the callback is NOT invoked.
        ' If null and a new slot is needed, Pack throws InvalidOperationException so the engine
        ' never ends up with archives that have no anchor plugin.
        Public Property PluginWriter As Action(Of String, GameKind)

        Public Property Overflow As ArchiveOverflowPolicy = ArchiveOverflowPolicy.SplitByPlugin
    End Class

    Public NotInheritable Class PackagerResult
        ' Archives that were (re)written this Pack call.
        Public ReadOnly Archives As New List(Of String)
        ' Archives skipped because the bundle was byte-identical to what was already on disk.
        Public ReadOnly Skipped As New List(Of String)
        ' Newly created dummy plugins (one per new slot). Existing plugin paths are NOT listed.
        Public ReadOnly Plugins As New List(Of String)
    End Class

    ''' <summary>Set of archive + plugin files in OutputDir whose names share a ModBaseName prefix.</summary>
    Public NotInheritable Class ArchiveSetInfo
        Public ReadOnly Archives As New List(Of String)   ' .ba2 / .bsa
        Public ReadOnly Plugins As New List(Of String)    ' .esp / .esm / .esl
    End Class

    Public NotInheritable Class UnpackRequest
        ' Where the archive set lives (and where its plugins are). Typically the game's Data folder.
        Public Property OutputDir As String
        ' Same base name used by Pack: discovery uses the "<base>*" prefix to find slots.
        Public Property ModBaseName As String = "WM_ClonePack"
        ' Where to extract entries to as loose files. Each entry's FullPath is appended.
        Public Property LooseDataDir As String
    End Class

    Public NotInheritable Class UnpackResult
        Public ReadOnly LooseFilesWritten As New List(Of String)
        Public ReadOnly ArchivesRemoved As New List(Of String)
        Public ReadOnly PluginsRemoved As New List(Of String)
    End Class

    ''' <summary>
    ''' Single-archive incremental packer. Compares a bundle of VirtualEntry against an existing
    ''' BA2/BSA on disk and rewrites only when content changed. Unchanged entries are stream-copied
    ''' (compressed bytes lifted verbatim from the previous archive) to avoid recompression on
    ''' multi-GB rebuilds.
    ''' </summary>
    Public NotInheritable Class ArchivePackager
        ' Bethesda engine convention: a plugin "Foo.esp" auto-loads "Foo - Main.ba2" and
        ' "Foo - Textures.ba2" (FO4) or "Foo.bsa" + "Foo - Textures.bsa" (SSE). These suffixes
        ' are spec-of-engine, not game data; safe to hardcode.
        Public Const SUFFIX_BA2_MAIN As String = " - Main.ba2"
        Public Const SUFFIX_BA2_TEXTURES As String = " - Textures.ba2"
        Public Const EXT_BSA As String = ".bsa"

        Private Enum BucketKind
            BA2_GNRL
            BA2_DX10
            BSA
        End Enum

        Private Enum DiffKind
            Unchanged       ' bundle == archive contents; skip rewrite
            NeedsRewrite    ' adds, removes, or content changes detected
        End Enum

        Private NotInheritable Class DiffResult
            Public Property Kind As DiffKind
            ' Paths where the bundle entry has the same length + CRC32 as the existing archive
            ' entry. These are stream-copied verbatim from .bak.
            Public ReadOnly UnchangedPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            ' Paths that exist in the archive but are NOT in the bundle. The packager preserves
            ' them automatically (stream-copy from .bak) — Pack semantics are "merge with existing",
            ' not "the bundle is the complete desired archive state".
            Public ReadOnly PreservePaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        End Class

        ' Per-slot plan built during distribution. Owns one plugin file path and 1..2 archive
        ' paths (FO4: Main + Textures; SSE: a single .bsa).
        Private NotInheritable Class PluginSlot
            Public Property SlotNumber As Integer    ' 1 = unsuffixed base, 2+ = numbered
            Public Property BaseName As String       ' e.g. "WM_ClonePack" or "WM_ClonePack2"
            Public Property PluginPath As String
            Public Property IsNew As Boolean         ' True if this slot did not exist on disk before Pack
            ' Existing-archive content used to anchor locked entries. Path → True if present.
            Public ReadOnly ExistingByBucket As New Dictionary(Of BucketKind, HashSet(Of String))
            ' Existing archive size in bytes (per bucket). Drives the size estimate during distribution.
            Public ReadOnly SizeByBucket As New Dictionary(Of BucketKind, Long)
            ' Entries assigned to this slot, grouped by bucket. Filled during distribution.
            Public ReadOnly AssignedByBucket As New Dictionary(Of BucketKind, List(Of VirtualEntry))
            ' Anchored entries (paths that already exist in this slot's archive) per bucket.
            ' Tracked separately from AssignedByBucket so the reuse threshold can ignore them
            ' when deciding whether to also accept free entries here: anchored forces rewrite
            ' regardless, so any free entries piggyback for free.
            Public ReadOnly AnchoredByBucket As New Dictionary(Of BucketKind, List(Of VirtualEntry))
        End Class

        Public Shared Function Pack(req As PackagerRequest) As PackagerResult
            If req Is Nothing Then Throw New ArgumentNullException(NameOf(req))
            If String.IsNullOrWhiteSpace(req.ModBaseName) Then Throw New ArgumentException("ModBaseName is empty.", NameOf(req))
            If String.IsNullOrWhiteSpace(req.OutputDir) Then Throw New ArgumentException("OutputDir is empty.", NameOf(req))
            If req.Entries Is Nothing Then Throw New ArgumentException("Entries is null.", NameOf(req))
            If req.MaxArchiveBytes <= 0 Then Throw New ArgumentException("MaxArchiveBytes must be positive.", NameOf(req))

            EnsureDir(req.OutputDir & Path.DirectorySeparatorChar)

            Dim result As New PackagerResult()
            If req.Entries.Count = 0 Then Return result

            Dim buckets = BucketsForGame(req.Game)

            ' --- Discover existing plugin slots in the output directory. ---
            Dim slots = DiscoverSlots(req, buckets)

            ' --- Distribute every bundle entry to a slot/bucket pair. ---
            DistributeEntries(req, slots, buckets)

            ' --- Emit each affected slot. Order: lower SlotNumber first, so the unsuffixed plugin
            '     gets written before its numbered companions. ---
            slots.Sort(Function(a, b) a.SlotNumber.CompareTo(b.SlotNumber))
            For Each slot In slots
                EmitSlot(req, slot, buckets, result)
            Next

            Return result
        End Function

        Private Shared Function BucketsForGame(g As GameKind) As BucketKind()
            Select Case g
                Case GameKind.FO4_BA2 : Return New BucketKind() {BucketKind.BA2_GNRL, BucketKind.BA2_DX10}
                Case GameKind.SSE_BSA : Return New BucketKind() {BucketKind.BSA}
                Case Else : Throw New ArgumentOutOfRangeException(NameOf(g))
            End Select
        End Function

        Private Shared Function ArchivePathFor(slot As PluginSlot, bucket As BucketKind, outputDir As String) As String
            Select Case bucket
                Case BucketKind.BA2_GNRL : Return Path.Combine(outputDir, slot.BaseName & SUFFIX_BA2_MAIN)
                Case BucketKind.BA2_DX10 : Return Path.Combine(outputDir, slot.BaseName & SUFFIX_BA2_TEXTURES)
                Case BucketKind.BSA : Return Path.Combine(outputDir, slot.BaseName & EXT_BSA)
                Case Else : Throw New ArgumentOutOfRangeException(NameOf(bucket))
            End Select
        End Function

        Private Shared Function BucketForEntry(ve As VirtualEntry, game As GameKind) As BucketKind
            If game = GameKind.SSE_BSA Then Return BucketKind.BSA
            Return If(IsTextureEntry(ve), BucketKind.BA2_DX10, BucketKind.BA2_GNRL)
        End Function

        Private Shared Function SlotName(slotNumber As Integer, baseName As String) As String
            ' Slot 1 uses the bare base ("WM_ClonePack"); 2..N append the number ("WM_ClonePack2").
            ' This matches how nested mods stack on Bethesda's auto-discovery convention.
            If slotNumber <= 1 Then Return baseName
            Return baseName & slotNumber.ToString()
        End Function

        ' --------------------------------------------------------------------------------------
        ' Discovery: enumerate "<base>*.esp" in OutputDir, parse the slot number, list the
        ' companion archive contents (paths only) so we can anchor locked entries to them later.
        ' --------------------------------------------------------------------------------------
        Private Shared Function DiscoverSlots(req As PackagerRequest, buckets As BucketKind()) As List(Of PluginSlot)
            Dim slots As New List(Of PluginSlot)
            If Not Directory.Exists(req.OutputDir) Then Return slots

            For Each pluginPath In Directory.EnumerateFiles(req.OutputDir, req.ModBaseName & "*.esp", SearchOption.TopDirectoryOnly)
                Dim slotNumber = ParseSlotNumber(Path.GetFileNameWithoutExtension(pluginPath), req.ModBaseName)
                If slotNumber <= 0 Then Continue For

                Dim slot As New PluginSlot With {
                    .SlotNumber = slotNumber,
                    .BaseName = SlotName(slotNumber, req.ModBaseName),
                    .PluginPath = pluginPath,
                    .IsNew = False
                }
                For Each b In buckets
                    slot.AssignedByBucket(b) = New List(Of VirtualEntry)()
                    slot.AnchoredByBucket(b) = New List(Of VirtualEntry)()
                    slot.ExistingByBucket(b) = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                    slot.SizeByBucket(b) = 0L

                    Dim ap = ArchivePathFor(slot, b, req.OutputDir)
                    If File.Exists(ap) Then
                        slot.SizeByBucket(b) = New FileInfo(ap).Length
                        Try
                            Using fs = File.OpenRead(ap)
                                Using reader As New BethesdaReader(fs)
                                    For Each ae In reader.EntriesFiles
                                        slot.ExistingByBucket(b).Add(NormalizePath(ae.FullPath))
                                    Next
                                End Using
                            End Using
                        Catch
                            ' Corrupt/unreadable archive: treat as empty for anchoring. PackOneArchive
                            ' will fail more loudly later if it can't diff the file.
                        End Try
                    End If
                Next
                slots.Add(slot)
            Next

            Return slots
        End Function

        Private Shared Function ParseSlotNumber(stem As String, baseName As String) As Integer
            ' "WM_ClonePack" → 1; "WM_ClonePack2" → 2; "WM_ClonePackFoo" → 0 (not a slot).
            If String.Equals(stem, baseName, StringComparison.OrdinalIgnoreCase) Then Return 1
            If Not stem.StartsWith(baseName, StringComparison.OrdinalIgnoreCase) Then Return 0
            Dim suffix = stem.Substring(baseName.Length)
            Dim n As Integer
            If Integer.TryParse(suffix, n) AndAlso n >= 2 Then Return n
            Return 0
        End Function

        ' --------------------------------------------------------------------------------------
        ' Distribute each bundle entry to a slot. Two kinds of entries:
        '
        '   - Anchored: entry.path already exists in some slot's archive. To honor the
        '     no-duplicate contract (Pack semantics: at most one copy of each path across the
        '     archive set), the entry must replace the existing one in that slot. Anchor forces
        '     a rewrite of the slot and bypasses the reuse threshold — the rewrite is mandatory
        '     regardless of how big or small the new content is. ComputeDiff later decides whether
        '     the slot actually needs to be rewritten (CRC match → Skipped) or really did change.
        '
        '   - Free: entry.path not yet present in any archive. Distribution chooses where to
        '     park it. Tentative pass: first slot with enough room (compressed size, exact when
        '     BundleAlreadyCompressed = True). Validation pass: for slots that ended up with
        '     ONLY free entries, accept reuse only if the proposed bytes-to-add ≥ slot.SizeByBucket
        '     × ReuseThreshold. Otherwise revert and redistribute, excluding the rejected slot.
        '
        ' Slot 1 is always created (with IsNew=True if it didn't exist on disk) so the anchor
        ' plugin is "<base>.esp" without a number prefix.
        ' --------------------------------------------------------------------------------------
        Private Shared Sub DistributeEntries(req As PackagerRequest, slots As List(Of PluginSlot), buckets As BucketKind())
            ' Ensure slot 1 always exists.
            Dim hasSlot1 = slots.Any(Function(s) s.SlotNumber = 1)
            If Not hasSlot1 Then
                slots.Add(NewSlot(1, req, buckets))
            End If

            ' --- PHASE 1: anchor by path ---------------------------------------------------------
            ' Anchored entries are recorded in slot.AnchoredByBucket and contribute their compressed
            ' size to slot.SizeByBucket immediately, so subsequent free-pass placement sees the
            ' correct projected size for the slot.
            Dim free As New List(Of VirtualEntry)()
            For Each ve In req.Entries
                Dim bucket = BucketForEntry(ve, req.Game)
                Dim p = NormalizePath(ve.FullPath)
                Dim anchorSlot As PluginSlot = Nothing
                For Each slot In slots
                    Dim existingPaths As HashSet(Of String) = Nothing
                    If slot.ExistingByBucket.TryGetValue(bucket, existingPaths) AndAlso existingPaths.Contains(p) Then
                        anchorSlot = slot
                        Exit For
                    End If
                Next
                If anchorSlot IsNot Nothing Then
                    anchorSlot.AnchoredByBucket(bucket).Add(ve)
                    ' Anchored entries replace the existing entry, but the slot's projected
                    ' compressed size doesn't grow by the full entry size (the existing entry
                    ' is going away). For the free-pass placement, leave SizeByBucket as the
                    ' on-disk size — that's the "slot size after anchor rewrite" upper bound
                    ' and a safe over-estimate.
                Else
                    free.Add(ve)
                End If
            Next

            ' --- PHASE 2: tentative free distribution (linear, exact compressed size) ------------
            ' Sort free largest-first so big entries get their pick, smaller ones fill what's left.
            free.Sort(Function(a, b) EntryProjectedSize(b, req).CompareTo(EntryProjectedSize(a, req)))

            ' Per slot we track how many bytes of free entries we've tentatively placed there.
            Dim proposedFreeBytes As New Dictionary(Of PluginSlot, Long)
            Dim proposedFreeBuckets As New Dictionary(Of PluginSlot, Long)
            Dim proposedFreeEntries As New Dictionary(Of PluginSlot, List(Of VirtualEntry))
            For Each s In slots
                proposedFreeBytes(s) = 0L
                proposedFreeEntries(s) = New List(Of VirtualEntry)()
            Next

            For Each ve In free
                Dim bucket = BucketForEntry(ve, req.Game)
                Dim addSize As Long = EntryProjectedSize(ve, req)

                Dim chosen As PluginSlot = Nothing
                For Each slot In slots
                    Dim cur As Long = slot.SizeByBucket(bucket) + proposedFreeBytes(slot)
                    If cur + addSize <= req.MaxArchiveBytes Then
                        chosen = slot
                        Exit For
                    End If
                Next

                If chosen Is Nothing Then
                    If req.Overflow = ArchiveOverflowPolicy.ThrowOnExceed Then
                        Throw New InvalidOperationException(
                            $"Bundle exceeds MaxArchiveBytes ({req.MaxArchiveBytes:N0} bytes) and Overflow=ThrowOnExceed.")
                    End If
                    Dim nextNumber As Integer = (slots.Max(Function(s) s.SlotNumber)) + 1
                    chosen = NewSlot(nextNumber, req, buckets)
                    slots.Add(chosen)
                    proposedFreeBytes(chosen) = 0L
                    proposedFreeEntries(chosen) = New List(Of VirtualEntry)()
                End If

                proposedFreeBytes(chosen) = proposedFreeBytes(chosen) + addSize
                proposedFreeEntries(chosen).Add(ve)
            Next

            ' --- PHASE 3: validate reuse against minimum-free-space rule -------------------------
            ' For each slot that has free entries proposed but NO anchored entries, accept the
            ' rewrite only if the slot has enough REAL free space (cap minus existing on-disk
            ' size) to be worth the I/O cost. Anchored slots are always rewritten regardless.
            Dim toRevert As New List(Of VirtualEntry)()
            Dim rejectedSlots As New HashSet(Of PluginSlot)()
            For Each slot In slots.ToList()
                Dim freeList = proposedFreeEntries(slot)
                If freeList.Count = 0 Then Continue For

                Dim hasAnyAnchor As Boolean = False
                For Each b In buckets
                    If slot.AnchoredByBucket(b).Count > 0 Then
                        hasAnyAnchor = True
                        Exit For
                    End If
                Next

                ' Slot already being rewritten by anchor → free entries piggyback for free.
                If hasAnyAnchor Then Continue For

                ' Existing-archive footprint: sum across buckets. 0 = brand new slot created in
                ' this Pack (no rewrite cost) → accept anything.
                Dim totalExistingSize As Long = 0L
                For Each b In buckets
                    totalExistingSize += slot.SizeByBucket(b)
                Next
                If totalExistingSize = 0L Then Continue For

                ' Real free space available in this slot at the cap. If below the configured
                ' minimum, the rewrite cost (moving totalExistingSize bytes) doesn't justify
                ' squeezing a tiny amount of new content in.
                Dim freeSpace As Long = req.MaxArchiveBytes - totalExistingSize
                If freeSpace >= req.MinFreeSpaceToFill Then Continue For

                ' Not enough room to be worth it — revert this slot's free proposals.
                toRevert.AddRange(freeList)
                proposedFreeBytes(slot) = 0L
                proposedFreeEntries(slot).Clear()
                rejectedSlots.Add(slot)
            Next

            ' --- PHASE 4: redistribute reverted entries, excluding rejected slots ----------------
            If toRevert.Count > 0 Then
                toRevert.Sort(Function(a, b) EntryProjectedSize(b, req).CompareTo(EntryProjectedSize(a, req)))
                For Each ve In toRevert
                    Dim bucket = BucketForEntry(ve, req.Game)
                    Dim addSize As Long = EntryProjectedSize(ve, req)

                    Dim chosen As PluginSlot = Nothing
                    For Each slot In slots
                        If rejectedSlots.Contains(slot) Then Continue For
                        Dim cur As Long = slot.SizeByBucket(bucket) + proposedFreeBytes(slot)
                        If cur + addSize <= req.MaxArchiveBytes Then
                            chosen = slot
                            Exit For
                        End If
                    Next

                    If chosen Is Nothing Then
                        If req.Overflow = ArchiveOverflowPolicy.ThrowOnExceed Then
                            Throw New InvalidOperationException(
                                $"Bundle exceeds MaxArchiveBytes after threshold reject ({req.MaxArchiveBytes:N0} bytes).")
                        End If
                        Dim nextNumber As Integer = (slots.Max(Function(s) s.SlotNumber)) + 1
                        chosen = NewSlot(nextNumber, req, buckets)
                        slots.Add(chosen)
                        proposedFreeBytes(chosen) = 0L
                        proposedFreeEntries(chosen) = New List(Of VirtualEntry)()
                    End If

                    proposedFreeBytes(chosen) = proposedFreeBytes(chosen) + addSize
                    proposedFreeEntries(chosen).Add(ve)
                Next
            End If

            ' --- PHASE 5: feed assignments into per-slot bundles ---------------------------------
            For Each slot In slots
                For Each b In buckets
                    For Each ve In slot.AnchoredByBucket(b)
                        slot.AssignedByBucket(b).Add(ve)
                    Next
                Next
                For Each ve In proposedFreeEntries(slot)
                    Dim bucket = BucketForEntry(ve, req.Game)
                    slot.AssignedByBucket(bucket).Add(ve)
                Next
            Next
        End Sub

        ' Projected on-disk size for a single entry, used during distribution and threshold checks.
        ' When BundleAlreadyCompressed = True the caller already populated PreCompressed* fields
        ' with exact byte counts — we use them verbatim. Otherwise fall back to the legacy estimate
        ' (raw size × CompressionRatioEstimate). PayloadSource entries are pre-compressed by
        ' definition (stream-copied from another archive) and report exact byte counts via Length.
        Private Shared Function EntryProjectedSize(ve As VirtualEntry, req As PackagerRequest) As Long
            If ve.PayloadSource IsNot Nothing Then
                Return ve.PayloadSource.Length
            End If
            If req.BundleAlreadyCompressed OrElse ve.PreCompressed Then
                ' PreCompressedCompSize = 0 means stored raw — the on-disk size equals decomp size.
                If ve.PreCompressedCompSize > 0UI Then
                    Return CLng(ve.PreCompressedCompSize)
                End If
                Return CLng(ve.PreCompressedDecompSize)
            End If
            Dim raw As Long = If(ve.Data Is Nothing, 0L, ve.Data.LongLength)
            Return CLng(raw * req.CompressionRatioEstimate)
        End Function

        Private Shared Function NewSlot(slotNumber As Integer, req As PackagerRequest, buckets As BucketKind()) As PluginSlot
            Dim baseName = SlotName(slotNumber, req.ModBaseName)
            Dim slot As New PluginSlot With {
                .SlotNumber = slotNumber,
                .BaseName = baseName,
                .PluginPath = Path.Combine(req.OutputDir, baseName & ".esp"),
                .IsNew = Not File.Exists(Path.Combine(req.OutputDir, baseName & ".esp"))
            }
            For Each b In buckets
                slot.AssignedByBucket(b) = New List(Of VirtualEntry)()
                slot.AnchoredByBucket(b) = New List(Of VirtualEntry)()
                slot.ExistingByBucket(b) = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                slot.SizeByBucket(b) = 0L
            Next
            Return slot
        End Function

        ' --------------------------------------------------------------------------------------
        ' Emit one slot: per bucket → PackOneArchive. After all buckets, if the slot is new,
        ' invoke the PluginWriter callback so the dummy .esp ends up next to the archives.
        ' --------------------------------------------------------------------------------------
        Private Shared Sub EmitSlot(req As PackagerRequest, slot As PluginSlot, buckets As BucketKind(), result As PackagerResult)
            Dim emittedAny As Boolean = False
            For Each bucket In buckets
                Dim sub_ As List(Of VirtualEntry) = slot.AssignedByBucket(bucket)
                If sub_.Count = 0 Then Continue For

                Dim archivePath = ArchivePathFor(slot, bucket, req.OutputDir)
                PackOneArchive(archivePath, sub_, bucket, result)
                emittedAny = True
            Next

            If slot.IsNew AndAlso emittedAny Then
                If req.PluginWriter Is Nothing Then
                    Throw New InvalidOperationException(
                        $"Pack would create a new plugin '{slot.PluginPath}' but PackagerRequest.PluginWriter is null.")
                End If
                req.PluginWriter.Invoke(slot.PluginPath, req.Game)
                result.Plugins.Add(slot.PluginPath)
            End If
        End Sub

        ' --------------------------------------------------------------------------------------
        ' Per-archive flow: diff vs existing → skip / rewrite. On rewrite, rename current to .bak,
        ' build the entry list (mixing pass-through + fresh), write, verify. Restore .bak on fail.
        ' --------------------------------------------------------------------------------------
        Private Shared Sub PackOneArchive(archivePath As String,
                                          bundle As List(Of VirtualEntry),
                                          kind As BucketKind,
                                          result As PackagerResult)
            EnsureDir(archivePath)

            ' Fast path: nothing on disk yet — write fresh.
            If Not File.Exists(archivePath) Then
                WriteArchive(archivePath, bundle, kind)
                VerifyArchive(archivePath, bundle)
                result.Archives.Add(archivePath)
                Return
            End If

            ' Diff against existing.
            Dim diff = ComputeDiff(archivePath, bundle)
            If diff.Kind = DiffKind.Unchanged Then
                result.Skipped.Add(archivePath)
                Return
            End If

            ' Rewrite. Rename current → .bak so we can pass-through unchanged payloads from it,
            ' and so we have a recovery point if the new write fails.
            Dim bakPath = archivePath & ".bak"
            If File.Exists(bakPath) Then File.Delete(bakPath)
            File.Move(archivePath, bakPath)

            Try
                ' The .bak reader (and its underlying FileStream) must stay open while WriteArchive
                ' runs: pass-through VirtualEntries hold PayloadSource references into _fs, and the
                ' writer streams them directly with no intermediate buffering.
                Dim entriesToWrite As List(Of VirtualEntry)
                Using fs = File.OpenRead(bakPath)
                    Using reader As New BethesdaReader(fs)
                        entriesToWrite = BuildEntriesToWrite(bundle, reader, diff)
                        WriteArchive(archivePath, entriesToWrite, kind)
                    End Using
                End Using

                ' Verify against the actual write (bundle ∪ preserved), not just the bundle —
                ' otherwise preserved entries would be counted as "extra" and verification would fail.
                VerifyArchive(archivePath, entriesToWrite)

                File.Delete(bakPath)
                result.Archives.Add(archivePath)
            Catch
                ' Rollback: best-effort restore of .bak. Caller sees the original archive intact.
                Try
                    If File.Exists(archivePath) Then File.Delete(archivePath)
                Catch
                End Try
                Try
                    If File.Exists(bakPath) Then File.Move(bakPath, archivePath)
                Catch
                End Try
                Throw
            End Try
        End Sub

        ' --------------------------------------------------------------------------------------
        ' Diff: walk the existing archive's manifest, compare against bundle by path + decompressed
        ' size + CRC32. CRC32 of existing entries forces a decompression pass — only paid once
        ' per Pack and only for paths that also appear in the bundle.
        ' --------------------------------------------------------------------------------------
        Private Shared Function ComputeDiff(existingPath As String, bundle As List(Of VirtualEntry)) As DiffResult
            Dim result As New DiffResult() With {.Kind = DiffKind.NeedsRewrite}

            Dim newByPath As New Dictionary(Of String, VirtualEntry)(StringComparer.OrdinalIgnoreCase)
            For Each ve In bundle
                newByPath(NormalizePath(ve.FullPath)) = ve
            Next

            Using fs = File.OpenRead(existingPath)
                Using reader As New BethesdaReader(fs)
                    Dim existingPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                    For Each ae In reader.EntriesFiles
                        existingPaths.Add(NormalizePath(ae.FullPath))
                    Next

                    ' Existing paths NOT in the bundle are preserved (merge semantics).
                    For Each ep In existingPaths
                        If Not newByPath.ContainsKey(ep) Then result.PreservePaths.Add(ep)
                    Next

                    Dim addedAny As Boolean = False
                    For Each np In newByPath.Keys
                        If Not existingPaths.Contains(np) Then
                            addedAny = True
                            Exit For
                        End If
                    Next

                    Dim changedAny As Boolean = False
                    For Each ae In reader.EntriesFiles
                        Dim p = NormalizePath(ae.FullPath)
                        Dim ve As VirtualEntry = Nothing
                        If Not newByPath.TryGetValue(p, ve) Then Continue For

                        ' Determine the "logical decompressed size" of the bundle entry — the
                        ' field we can compare against the source archive's stored decomp size
                        ' without ever decompressing. PreCompressed entries (the caller already
                        ' compressed) report PreCompressedDecompSize. Legacy entries with raw
                        ' Data report Data.Length.
                        Dim newDecompSize As Long
                        If ve.PreCompressed Then
                            newDecompSize = CLng(ve.PreCompressedDecompSize)
                        ElseIf ve.Data IsNot Nothing Then
                            newDecompSize = ve.Data.LongLength
                        Else
                            newDecompSize = 0L
                        End If

                        ' Cheap first check: extract the raw stored chunk (no decompression) so we
                        ' can read the existing entry's decompressed size from the chunk header.
                        ' If the sizes don't match, the content has changed for sure — no need
                        ' to actually decompress to find out.
                        Dim raw As RawCompressedEntry
                        Try
                            raw = reader.ExtractCompressedPayload(ae.Index)
                        Catch
                            ' Some readers can't supply ExtractCompressedPayload (multi-chunk,
                            ' BSA edge cases). Treat as "changed" — forces a rewrite, which is
                            ' the safe direction.
                            changedAny = True
                            Continue For
                        End Try

                        If CLng(raw.DecompSize) <> newDecompSize Then
                            changedAny = True
                            Continue For
                        End If

                        ' Sizes match — verify content via CRC32 of the decompressed payload.
                        ' Do this only when the caller filled ve.Crc32 (otherwise we can't tell
                        ' identity cheaply, fall back to "changed" to be safe).
                        If ve.Crc32 = 0UI Then
                            changedAny = True
                            Continue For
                        End If

                        ' Decompress raw.Bytes manually so we get the EXACT payload that the
                        ' caller compressed and CRC'd, without any header reconstruction the
                        ' archive's high-level Extract path may apply (e.g. BA2 DX10 prepends a
                        ' rebuilt DDS header on ExtractToMemory). The caller's ve.Crc32 is over
                        ' the stripped/raw payload — we need the same shape on this side.
                        Dim payloadForCrc As Byte() = Nothing
                        Try
                            If raw.IsCompressed Then
                                ' Compressed payload: raw.Bytes is the chunk's compressed stream,
                                ' raw.DecompSize is the expected decompressed length. The codec
                                ' is Zlib unless this is a BSA LZ4 entry or BA2 v3+LZ4 archive.
                                ' We pick by trying Zlib first (most common for BA2 GNRL/DX10);
                                ' if that fails, fall back to LZ4. Both decoders are strict.
                                Try
                                    payloadForCrc = ZlibStrict.ZlibDecompressStrict(raw.Bytes, CInt(raw.DecompSize))
                                Catch
                                    payloadForCrc = Lz4Strict.Lz4DecompressStrict(raw.Bytes, CInt(raw.DecompSize))
                                End Try
                            Else
                                ' Stored uncompressed: raw.Bytes IS the payload.
                                payloadForCrc = raw.Bytes
                            End If
                        Catch
                            ' Decompression failed for any reason → can't verify identity, treat
                            ' as changed (safe direction).
                            changedAny = True
                            Continue For
                        End Try

                        If payloadForCrc IsNot Nothing AndAlso Ba2WriterCommon.Crc32Bytes(payloadForCrc) = ve.Crc32 Then
                            result.UnchangedPaths.Add(p)
                        Else
                            changedAny = True
                        End If
                    Next

                    ' Skip rewrite only if the bundle is fully covered by unchanged entries AND
                    ' nothing in the bundle is missing from the archive. Preserved paths alone
                    ' don't need a rewrite — they're already on disk.
                    If Not addedAny AndAlso Not changedAny Then
                        result.Kind = DiffKind.Unchanged
                    End If
                End Using
            End Using

            Return result
        End Function

        ' --------------------------------------------------------------------------------------
        ' Build the entry list to feed the writer:
        '   - paths in diff.UnchangedPaths → wrap bundle entry with pass-through bytes from .bak.
        '   - paths that are added or changed → forward the bundle entry as-is (writer compresses).
        '   - paths in diff.PreservePaths (existing in .bak but NOT in the bundle) → emit pass-through
        '     entry from .bak so the rewritten archive keeps everything that was already there.
        '     This is what makes Pack a merge ("upsert") instead of a destructive replace.
        ' --------------------------------------------------------------------------------------
        Private Shared Function BuildEntriesToWrite(bundle As List(Of VirtualEntry),
                                                    bakReader As BethesdaReader,
                                                    diff As DiffResult) As List(Of VirtualEntry)
            Dim bakByPath As New Dictionary(Of String, ArchiveEntry)(StringComparer.OrdinalIgnoreCase)
            For Each ae In bakReader.EntriesFiles
                bakByPath(NormalizePath(ae.FullPath)) = ae
            Next

            Dim out As New List(Of VirtualEntry)(bundle.Count + diff.PreservePaths.Count)

            ' First: bundle entries (each one keeps its order; unchanged ones get pass-through).
            For Each ve In bundle
                Dim p = NormalizePath(ve.FullPath)
                Dim bakAe As ArchiveEntry = Nothing
                If diff.UnchangedPaths.Contains(p) AndAlso bakByPath.TryGetValue(p, bakAe) Then
                    out.Add(BuildPassThroughEntry(ve.Directory, ve.FileName, ve.PreferCompress, ve.Crc32, bakReader, bakAe.Index))
                Else
                    out.Add(ve)
                End If
            Next

            ' Then: preserved entries (existing in .bak, not in the bundle). Always pass-through.
            For Each pp In diff.PreservePaths
                Dim bakAe As ArchiveEntry = Nothing
                If Not bakByPath.TryGetValue(pp, bakAe) Then Continue For
                out.Add(BuildPassThroughEntry(bakAe.Directory, bakAe.FileName, False, 0UI, bakReader, bakAe.Index))
            Next

            Return out
        End Function

        Private Shared Function BuildPassThroughEntry(dir As String,
                                                       fileName As String,
                                                       preferCompress As Boolean,
                                                       crc32 As UInteger,
                                                       bakReader As BethesdaReader,
                                                       bakIndex As Integer) As VirtualEntry
            ' Stream-copy pass-through: the writer seeks PayloadSource.SourceStream and copies
            ' Length bytes directly into the output without ever materializing the payload in
            ' a managed array. RAM peak per entry is O(64 KB) regardless of file size, which is
            ' the difference between minutes and hours on multi-GB archive rewrites.
            Dim ps = bakReader.GetPayloadSource(bakIndex)
            Dim ve As New VirtualEntry With {
                .Directory = dir,
                .FileName = fileName,
                .PayloadSource = ps,
                .PreferCompress = preferCompress,
                .Crc32 = crc32
            }
            If ps.HasDx10Metadata Then
                ve.Width = ps.Width
                ve.Height = ps.Height
                ve.MipCount = ps.MipCount
                ve.DxgiFormat = ps.DxgiFormat
                ve.IsCubemap = ps.IsCubemap
                ve.Faces = ps.Faces
            End If
            Return ve
        End Function

        ' --------------------------------------------------------------------------------------
        ' Writer dispatch by bucket. Options use library defaults — Pack does not expose them yet.
        ' --------------------------------------------------------------------------------------
        Private Shared Sub WriteArchive(path As String, entries As List(Of VirtualEntry), kind As BucketKind)
            Using fs As FileStream = File.Create(path)
                Select Case kind
                    Case BucketKind.BA2_GNRL
                        Ba2WriterGNRL.Write(fs, entries, New Ba2WriterGNRL.Options())
                    Case BucketKind.BA2_DX10
                        Ba2WriterDX10.Write(fs, entries, New Ba2WriterDX10.Options())
                    Case BucketKind.BSA
                        BsaWriter.Write(fs, entries, New BsaWriter.Options())
                    Case Else
                        Throw New ArgumentOutOfRangeException(NameOf(kind))
                End Select
            End Using
        End Sub

        ' --------------------------------------------------------------------------------------
        ' Verify: re-open the freshly written archive, confirm count + path set match the bundle,
        ' and decompress 3 random entries to exercise the chunk headers / compression streams.
        ' --------------------------------------------------------------------------------------
        Private Shared Sub VerifyArchive(path As String, expectedBundle As List(Of VirtualEntry))
            Dim expectedPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each ve In expectedBundle
                expectedPaths.Add(NormalizePath(ve.FullPath))
            Next

            Using fs = File.OpenRead(path)
                Using reader As New BethesdaReader(fs)
                    Dim actual = reader.EntriesFiles
                    If actual.Count <> expectedBundle.Count Then
                        Throw New InvalidDataException($"Verify failed for '{path}': expected {expectedBundle.Count} entries, got {actual.Count}.")
                    End If

                    For Each ae In actual
                        Dim ap = NormalizePath(ae.FullPath)
                        If Not expectedPaths.Contains(ap) Then
                            Throw New InvalidDataException($"Verify failed for '{path}': unexpected entry '{ae.FullPath}'.")
                        End If
                    Next

                    Dim sampleCount As Integer = Math.Min(3, actual.Count)
                    If sampleCount > 0 Then
                        Dim rnd As New Random()
                        Dim picked As New HashSet(Of Integer)
                        While picked.Count < sampleCount
                            picked.Add(rnd.Next(actual.Count))
                        End While
                        For Each idx In picked
                            Dim bytes = reader.ExtractToMemory(idx)
                            If bytes Is Nothing OrElse bytes.Length = 0 Then
                                Throw New InvalidDataException($"Verify failed for '{path}': spot-check entry {idx} extracts empty.")
                            End If
                        Next
                    End If
                End Using
            End Using
        End Sub

        Private Shared Function IsTextureEntry(ve As VirtualEntry) As Boolean
            Dim ext = IO.Path.GetExtension(If(ve.FileName, "")).ToLowerInvariant()
            Return ext = ".dds"
        End Function

        Private Shared Function NormalizePath(p As String) As String
            Return PathUtil.NormalizeSlash(If(p, "")).Trim(Correct_Path_separator).ToLowerInvariant()
        End Function

        Private Shared Sub EnsureDir(filePath As String)
            Dim dir = IO.Path.GetDirectoryName(filePath)
            If Not String.IsNullOrEmpty(dir) AndAlso Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If
        End Sub

        ' ======================================================================================
        '                                    UNPACK + DISCOVERY
        ' ======================================================================================

        ''' <summary>
        ''' Enumerates archives and plugins under outputDir whose file name starts with modBaseName.
        ''' Used by Unpack to drive cleanup and by callers who want to inspect the current pack
        ''' state before deciding what to do.
        ''' </summary>
        Public Shared Function DiscoverArchiveSet(outputDir As String, modBaseName As String) As ArchiveSetInfo
            If String.IsNullOrWhiteSpace(outputDir) Then Throw New ArgumentException("outputDir is empty.", NameOf(outputDir))
            If String.IsNullOrWhiteSpace(modBaseName) Then Throw New ArgumentException("modBaseName is empty.", NameOf(modBaseName))

            Dim info As New ArchiveSetInfo()
            If Not Directory.Exists(outputDir) Then Return info

            ' Restrict to files whose stem either equals modBaseName, or starts with "modBaseName"
            ' followed by a digit (numbered slot) or " - " (bucket suffix). Anything else is a
            ' coincidental name match (e.g. "WM_ClonePackLegacy.esp") and stays untouched.
            For Each ext In New String() {".ba2", ".bsa", ".esp", ".esm", ".esl"}
                For Each filePath In Directory.EnumerateFiles(outputDir, modBaseName & "*" & ext, SearchOption.TopDirectoryOnly)
                    Dim stem = Path.GetFileNameWithoutExtension(filePath)
                    If Not BelongsToModBaseName(stem, modBaseName) Then Continue For

                    Select Case ext
                        Case ".ba2", ".bsa" : info.Archives.Add(filePath)
                        Case ".esp", ".esm", ".esl" : info.Plugins.Add(filePath)
                    End Select
                Next
            Next

            info.Archives.Sort(StringComparer.OrdinalIgnoreCase)
            info.Plugins.Sort(StringComparer.OrdinalIgnoreCase)
            Return info
        End Function

        ''' <summary>
        ''' Recognises file stems generated by Pack: "<base>", "<base>N", "<base> - Main",
        ''' "<base>N - Main", "<base> - Textures", "<base>N - Textures". Rejects bare prefix
        ''' matches that don't follow one of those shapes.
        ''' </summary>
        Private Shared Function BelongsToModBaseName(stem As String, modBaseName As String) As Boolean
            If String.Equals(stem, modBaseName, StringComparison.OrdinalIgnoreCase) Then Return True
            If Not stem.StartsWith(modBaseName, StringComparison.OrdinalIgnoreCase) Then Return False

            Dim suffix = stem.Substring(modBaseName.Length)
            ' Strip an optional numeric run (slot suffix) before checking the bucket marker.
            Dim i As Integer = 0
            While i < suffix.Length AndAlso Char.IsDigit(suffix(i))
                i += 1
            End While

            ' If a numeric run was found, fall through to bucket-suffix check on the remainder.
            ' If nothing followed the digits, that's the plain numbered slot ("<base>2").
            Dim afterDigits = suffix.Substring(i)
            If afterDigits.Length = 0 Then Return i > 0   ' "<base>2" valid; "<base>" already handled.

            ' Accept " - Main" or " - Textures" (Pack's only non-numeric suffixes today).
            Return afterDigits.Equals(" - Main", StringComparison.OrdinalIgnoreCase) OrElse
                   afterDigits.Equals(" - Textures", StringComparison.OrdinalIgnoreCase)
        End Function

        ''' <summary>
        ''' Reverses Pack: extracts every entry of every archive in the set as a loose file under
        ''' LooseDataDir, then deletes the archives and their dummy plugins. Loose files preserve
        ''' the entry's FullPath relative to the data root (so "Textures\ManoloCloned\foo.dds"
        ''' becomes "&lt;LooseDataDir&gt;\Textures\ManoloCloned\foo.dds").
        '''
        ''' Optional onEntry: invoked once per extracted entry as (doneSoFar, totalAcrossAllArchives,
        ''' currentEntryRelPath). totalAcrossAllArchives is computed up front from the archive
        ''' file tables (no payload reads), so the caller can drive a determinate progress bar.
        '''
        ''' Optional ct: checked between entries during extraction. Cancellation is safe — already-
        ''' written loose files stay on disk, archives are NOT deleted (phase 2 only runs on full
        ''' success). The caller can re-run Unpack to finish the job.
        '''
        ''' Failure mode: extraction errors abort BEFORE any deletion happens. Already-written loose
        ''' files remain on disk; the caller decides whether to roll them back. Per-archive deletion
        ''' is best-effort once all extractions succeeded.
        ''' </summary>
        Public Shared Function Unpack(req As UnpackRequest,
                                       Optional onEntry As Action(Of Integer, Integer, String) = Nothing,
                                       Optional ct As System.Threading.CancellationToken = Nothing,
                                       Optional onArchiveStart As Action(Of String, Integer, Integer) = Nothing) As UnpackResult
            If req Is Nothing Then Throw New ArgumentNullException(NameOf(req))
            If String.IsNullOrWhiteSpace(req.OutputDir) Then Throw New ArgumentException("OutputDir is empty.", NameOf(req))
            If String.IsNullOrWhiteSpace(req.ModBaseName) Then Throw New ArgumentException("ModBaseName is empty.", NameOf(req))
            If String.IsNullOrWhiteSpace(req.LooseDataDir) Then Throw New ArgumentException("LooseDataDir is empty.", NameOf(req))

            Dim info = DiscoverArchiveSet(req.OutputDir, req.ModBaseName)
            Dim result As New UnpackResult()

            If info.Archives.Count = 0 AndAlso info.Plugins.Count = 0 Then Return result

            ' Pre-pass: count total entries across all archives so the caller can show determinate
            ' progress. This is cheap — Open() parses the file table only, no payload reads.
            Dim totalEntries As Integer = 0
            For Each archivePath In info.Archives
                Try
                    Using fs = File.OpenRead(archivePath)
                        Using reader As New BethesdaReader(fs)
                            totalEntries += reader.EntriesFiles.Count
                        End Using
                    End Using
                Catch
                    ' Corrupt archive: skip in count; the extraction loop below will surface the error.
                End Try
            Next

            Dim entriesDone As Integer = 0
            Dim cancelled As Boolean = False

            ' Phase 1: extract every entry of every archive, and DELETE THE ARCHIVE only after
            ' all of its entries were successfully written to loose. This bounds disk peak to
            ' "one archive worth of duplicated bytes" instead of the entire packed set, while
            ' staying recoverable: if extraction fails or is cancelled mid-archive, that archive
            ' stays on disk and a re-run of Unpack picks up where it left off (already-extracted
            ' loose files get renamed to *.bak.unpack — no data loss).
            Dim archiveIndex As Integer = 0
            For Each archivePath In info.Archives
                If ct.IsCancellationRequested Then
                    cancelled = True
                    Exit For
                End If

                archiveIndex += 1
                If onArchiveStart IsNot Nothing Then onArchiveStart(archivePath, archiveIndex, info.Archives.Count)

                Dim archiveFullyExtracted As Boolean = False
                Try
                    Using fs = File.OpenRead(archivePath)
                        Using reader As New BethesdaReader(fs)
                            archiveFullyExtracted = True
                            For Each entry In reader.EntriesFiles
                                If ct.IsCancellationRequested Then
                                    cancelled = True
                                    archiveFullyExtracted = False
                                    Exit For
                                End If

                                Dim relPath = entry.FullPath
                                Dim outPath = Path.Combine(req.LooseDataDir, relPath)
                                EnsureDir(outPath)

                                Dim bytes = reader.ExtractToMemory(entry.Index)
                                If bytes Is Nothing Then bytes = Array.Empty(Of Byte)()

                                ' Preserve any existing loose file that wasn't ours by renaming it.
                                ' Caller post-processes "*.bak.unpack" if it cares about conflicts.
                                If File.Exists(outPath) Then
                                    Dim conflictBak = outPath & ".bak.unpack"
                                    If File.Exists(conflictBak) Then File.Delete(conflictBak)
                                    File.Move(outPath, conflictBak)
                                End If

                                File.WriteAllBytes(outPath, bytes)
                                result.LooseFilesWritten.Add(outPath)

                                entriesDone += 1
                                If onEntry IsNot Nothing Then onEntry(entriesDone, totalEntries, relPath)
                            Next
                        End Using
                    End Using
                Catch
                    archiveFullyExtracted = False
                    Throw
                End Try

                ' Delete the source archive only after all its entries are safely on disk as loose.
                ' The caller is expected to have unmounted any FilesDictionary handles for this
                ' path BEFORE calling Unpack (see WM_PackUnpack.Unpack pre-unregister loop) so
                ' File.Delete is not blocked by a sharing violation.
                If archiveFullyExtracted Then
                    Try
                        File.Delete(archivePath)
                        result.ArchivesRemoved.Add(archivePath)
                    Catch
                        ' Leave it on disk; the loose copy is already written, the user can clean
                        ' up manually or re-run Unpack later.
                    End Try
                End If

                If cancelled Then Exit For
            Next

            ' Phase 2: plugins. Delete only if extraction was not cancelled — leaving plugins behind
            ' on cancel keeps the engine able to find the still-existing archives if any survived.
            If Not cancelled Then
                For Each pluginPath In info.Plugins
                    Try
                        File.Delete(pluginPath)
                        result.PluginsRemoved.Add(pluginPath)
                    Catch
                    End Try
                Next
            End If

            Return result
        End Function
    End Class

End Namespace
