using stdole;

public class IPictureDispConverter : AxHost
{
    private IPictureDispConverter() : base("")
    {
    }

    public static Image? PictureDispToImage(IPictureDisp pictureDisp)
    {
        try
        {
            return GetPictureFromIPicture(pictureDisp);
        }
        catch
        {
            return null;
        }
    }
}