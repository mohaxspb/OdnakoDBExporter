/**
 * 
 */
package ru.kuchanov.databaseexporter.db;

import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import android.content.Context;
import android.content.Intent;
import android.database.Cursor;
import android.net.Uri;
import android.os.AsyncTask;
import android.os.Environment;
import android.util.Log;

/**
 * @author Юрий
 * 
 */
public class XlsWriter extends AsyncTask<Void, Void, File[]>
{
	final private static String LOG = XlsWriter.class.getSimpleName();

	String to = "mohax.spb@gmail.com";
	String subj = "odnako, db, ormlite, table, csv, xls, excel, java, android";
	String msg = "test";

	private Context ctx;
	private Cursor c;

	public XlsWriter(Context ctx, Cursor c)
	{
		//		Log.e(LOG, "constructor");
		this.ctx = ctx;
		this.c = c;
	}

	@Override
	protected File[] doInBackground(Void... params)
	{
		//		File dbFile = this.ctx.getDatabasePath(DataBaseHelper.DATABASE_NAME);
		//		Log.v(LOG, "Db path is: " + dbFile); //get the path of db

		//TODO SIZE!
		File[] output = new File[1];

		output[0] = this.getArticleXLS();
		//		output[1] = this.getCategoryCSV();
		//		output[2] = this.getArtCatTableCSV();
		//		output[3] = this.getAuthorCSV();
		//		output[4] = this.getArtAutTableCSV();

		return output;
	}

	@Override
	protected void onPostExecute(File[] out)
	{
		Log.e(LOG, "onPostExecute");
		//		Intent emailIntent = new Intent(android.content.Intent.ACTION_SEND);
		Intent emailIntent = new Intent(Intent.ACTION_SEND_MULTIPLE);
		emailIntent.setType("text/html");
		emailIntent.putExtra(android.content.Intent.EXTRA_EMAIL, new String[] { to });
		emailIntent.putExtra(android.content.Intent.EXTRA_SUBJECT, subj);
		emailIntent.putExtra(android.content.Intent.EXTRA_TEXT, msg);

		ArrayList<Uri> uris = new ArrayList<Uri>();
		for (File f : out)
		{
			//			emailIntent.putExtra(Intent.EXTRA_STREAM, Uri.fromFile(f));
			uris.add(Uri.fromFile(f));
		}
		emailIntent.putParcelableArrayListExtra(Intent.EXTRA_STREAM, uris);
		//		this.ctx.startActivityForResult(Intent.createChooser(emailIntent, "Sending multiple attachment"), 12345);

		//				emailIntent.putExtra(Intent.EXTRA_STREAM, Uri.fromFile(file));
		this.ctx.startActivity(Intent.createChooser(emailIntent, "Send mail..."));

	}

	//	private File getArtAutTableCSV()
	//	{
	//	
	//	}
	//
	//	private File getAuthorCSV()
	//	{
	//		
	//	}
	//
	//	private File getArtCatTableCSV()
	//	{
	//		
	//	}
	//
	//	private File getCategoryCSV()
	//	{
	//		
	//	}

	private File getArticleXLS()
	{
		try
		{
			//			File exportDir = new File(this.ctx.getExternalFilesDir(null), "");
			File exportDir = new File(Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS),
			"");

			String path = exportDir.getCanonicalPath();

			//Instantiate a Workbook object.
			Workbook workbook = new Workbook();
			//Get the first worksheet's cells in the book.
			Cells cells = workbook.getWorksheets().get(0).getCells();

			//get column names from DB table
//			String[] columnNames = new String[c.getColumnCount()];
//
//			for (int i = 0; i < c.getColumnCount(); i++)
//			{
//				columnNames[i] = c.getColumnName(i);
//				Log.e(LOG, columnNames[i]);
//			}
			String[] columnNames = getArticleFieldsNames();

			for (int i = 0; i < columnNames.length; i++)
			{
				Log.e(LOG, columnNames[i]);
			}

			//insert table data
			if (c.moveToFirst())
			{
				while (!c.isAfterLast())
				{
					if (c.getPosition() == 0)
					{
						//insert column names
						for (int i = 0; i < columnNames.length; i++)
						{
							cells.get(0, i).setValue(columnNames[i]);
						}
					}
					else
					{
						for (int i = 0; i < c.getColumnCount(); i++)
						{
							String data = c.getString(c.getColumnIndex(columnNames[i]));
							cells.get(c.getPosition(), i).setValue(data);
						}
					}
					c.moveToNext();
				}
			}

			//Save the Excel file.
			workbook.save(path + "/ArticleXLS.xls");

			File file = new File(path + "/ArticleXLS.xls");

			return file;

		} catch (IOException e)
		{
			Log.e(LOG, e.getMessage());
			return null;
		} catch (Exception e)
		{
			Log.e(LOG, e.getMessage());
			return null;
		}
	}
	
	/**
	 * returns String arr with names of all Table columns
	 */
	public static String[] getArticleFieldsNames()
	{
		String[] arrStr1 = { "id", "url", "title", "img_art", "authorBlogUrl",
				"authorName", "preview", "pubDate", "refreshed", "numOfComments",
				"numOfSharings", "artText", "authorDescr", "tegs_main", "tegs_all",
				"share_quont", "to_read_main", "to_read_more", "img_author", "author" };
		return arrStr1;
	}
}
