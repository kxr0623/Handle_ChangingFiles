import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class HSSFWorkbook {
    POIFSFileSystem pfs;
    public HSSFWorkbook(POIFSFileSystem pfs) {
        this.pfs=pfs;
    }
}
