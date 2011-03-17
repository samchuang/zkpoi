
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


package org.zkoss.poi.poifs.filesystem;

import java.io.IOException;
import java.nio.ByteBuffer;
import java.util.Iterator;
import java.util.List;

import org.zkoss.poi.poifs.common.POIFSConstants;
import org.zkoss.poi.poifs.property.RootProperty;
import org.zkoss.poi.poifs.storage.BATBlock;
import org.zkoss.poi.poifs.storage.BlockAllocationTableWriter;
import org.zkoss.poi.poifs.storage.HeaderBlock;
import org.zkoss.poi.poifs.storage.BATBlock.BATBlockAndIndex;

/**
 * This class handles the MiniStream (small block store)
 *  in the NIO case for {@link NPOIFSFileSystem}
 */
public class NPOIFSMiniStore extends BlockStore
{
    private NPOIFSFileSystem _filesystem;
    private NPOIFSStream     _mini_stream;
    private List<BATBlock>   _sbat_blocks;
    private HeaderBlock      _header;
    private RootProperty     _root;

    protected NPOIFSMiniStore(NPOIFSFileSystem filesystem, RootProperty root,
         List<BATBlock> sbats, HeaderBlock header)
    {
       this._filesystem = filesystem;
       this._sbat_blocks = sbats;
       this._header = header;
       this._root = root;
       
       this._mini_stream = new NPOIFSStream(filesystem, root.getStartBlock());
    }
    
    /**
     * Load the block at the given offset.
     */
    protected ByteBuffer getBlockAt(final int offset) throws IOException {
       // Which big block is this?
       int byteOffset = offset * POIFSConstants.SMALL_BLOCK_SIZE;
       int bigBlockNumber = byteOffset / _filesystem.getBigBlockSize();
       int bigBlockOffset = byteOffset % _filesystem.getBigBlockSize();
       
       // Now locate the data block for it
       Iterator<ByteBuffer> it = _mini_stream.getBlockIterator();
       for(int i=0; i<bigBlockNumber; i++) {
          it.next();
       }
       ByteBuffer dataBlock = it.next();
       
       // Our blocks are small, so duplicating it is fine 
       byte[] data = new byte[POIFSConstants.SMALL_BLOCK_SIZE];
       dataBlock.position(
             dataBlock.position() + bigBlockOffset
       );
       dataBlock.get(data, 0, data.length);
       
       // Return a ByteBuffer on this
       ByteBuffer miniBuffer = ByteBuffer.wrap(data);
       return miniBuffer;
    }
    
    /**
     * Load the block, extending the underlying stream if needed
     */
    protected ByteBuffer createBlockIfNeeded(final int offset) throws IOException {
       // TODO Extend the stream if needed
       // TODO Needs append support on the underlying stream
       return getBlockAt(offset);
    }
    
    /**
     * Returns the BATBlock that handles the specified offset,
     *  and the relative index within it
     */
    protected BATBlockAndIndex getBATBlockAndIndex(final int offset) {
       return BATBlock.getSBATBlockAndIndex(
             offset, _header, _sbat_blocks
       );
    }
    
    /**
     * Works out what block follows the specified one.
     */
    protected int getNextBlock(final int offset) {
       BATBlockAndIndex bai = getBATBlockAndIndex(offset);
       return bai.getBlock().getValueAt( bai.getIndex() );
    }
    
    /**
     * Changes the record of what block follows the specified one.
     */
    protected void setNextBlock(final int offset, final int nextBlock) {
       BATBlockAndIndex bai = getBATBlockAndIndex(offset);
       bai.getBlock().setValueAt(
             bai.getIndex(), nextBlock
       );
    }
    
    /**
     * Finds a free block, and returns its offset.
     * This method will extend the file if needed, and if doing
     *  so, allocate new FAT blocks to address the extra space.
     */
    protected int getFreeBlock() throws IOException {
       int sectorsPerSBAT = _filesystem.getBigBlockSizeDetails().getBATEntriesPerBlock();
       
       // First up, do we have any spare ones?
       int offset = 0;
       for(int i=0; i<_sbat_blocks.size(); i++) {
          // Check this one
          BATBlock sbat = _sbat_blocks.get(i);
          if(sbat.hasFreeSectors()) {
             // Claim one of them and return it
             for(int j=0; j<sectorsPerSBAT; j++) {
                int sbatValue = sbat.getValueAt(j);
                if(sbatValue == POIFSConstants.UNUSED_BLOCK) {
                   // Bingo
                   return offset + j;
                }
             }
          }
          
          // Move onto the next SBAT
          offset += sectorsPerSBAT;
       }
       
       // If we get here, then there aren't any
       //  free sectors in any of the SBATs
       // So, we need to extend the chain and add another
       
       // Create a new BATBlock
       BATBlock newSBAT = BATBlock.createEmptyBATBlock(_filesystem.getBigBlockSizeDetails(), false);
       int batForSBAT = _filesystem.getFreeBlock();
       newSBAT.setOurBlockIndex(batForSBAT);
       
       // Are we the first SBAT?
       if(_header.getSBATCount() == 0) {
          _header.setSBATStart(batForSBAT);
          _header.setSBATBlockCount(1);
       } else {
          // Find the end of the SBAT stream, and add the sbat in there
          ChainLoopDetector loopDetector = _filesystem.getChainLoopDetector();
          int batOffset = _header.getSBATStart();
          while(true) {
             loopDetector.claim(batOffset);
             int nextBat = _filesystem.getNextBlock(batOffset);
             if(nextBat == POIFSConstants.END_OF_CHAIN) {
                break;
             }
             batOffset = nextBat;
          }
          
          // Add it in at the end
          _filesystem.setNextBlock(batOffset, batForSBAT);
          
          // And update the count
          _header.setSBATBlockCount(
                _header.getSBATCount() + 1
          );
       }
       
       // Finish allocating
       _filesystem.setNextBlock(batForSBAT, POIFSConstants.END_OF_CHAIN);
       _sbat_blocks.add(newSBAT);
       
       // Return our first spot
       return offset;
    }
    
    @Override
    protected ChainLoopDetector getChainLoopDetector() throws IOException {
      return new ChainLoopDetector( _root.getSize() );
    }

    protected int getBlockStoreBlockSize() {
       return POIFSConstants.SMALL_BLOCK_SIZE;
    }
    
    /**
     * Writes the SBATs to their backing blocks
     */
    protected void syncWithDataSource() throws IOException {
       for(BATBlock sbat : _sbat_blocks) {
          ByteBuffer block = _filesystem.getBlockAt(sbat.getOurBlockIndex());
          BlockAllocationTableWriter.writeBlock(sbat, block);
       }
    }
}
