/**
 * 
 */
package ru.kuchanov.databaseexporter.db;

import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

import android.content.ContentProviderClient;
import android.content.Context;
import android.content.Intent;
import android.database.Cursor;
import android.net.Uri;
import android.os.AsyncTask;
import android.os.Environment;
import android.os.RemoteException;
import android.util.Log;

/**
 * @author Юрий
 * 
 */
public class XlsWriterTask extends AsyncTask<Void, Void, File[]>
{
	final private static String LOG = XlsWriterTask.class.getSimpleName();

	String to = "mohax.spb@gmail.com";
	String subj = "odnako, db, ormlite, table, csv, xls, excel, java, android";
	String msg = "test";

	private Context ctx;

	public XlsWriterTask(Context ctx)
	{
		this.ctx = ctx;
	}

	@Override
	protected File[] doInBackground(Void... cursors)
	{
		//TODO SIZE!
		File[] output = new File[3];

		//get article table and write it to xls
		Uri articleAllURI = Uri.parse("content://ru.kuchanov.odnako.db.ContentProviderOdnakoDB/article");
		ContentProviderClient yourCR = ctx.getContentResolver().acquireContentProviderClient(articleAllURI);
		try
		{
			Cursor articleCursor = yourCR.query(articleAllURI, null, null, null, null);
			output[0] = this.getArticleXLS(articleCursor);
		} catch (RemoteException e)
		{
			e.printStackTrace();
		}

		//get ArtCatTable table and write it to xls
		Uri artCatAllURI = Uri.parse("content://ru.kuchanov.odnako.db.ContentProviderOdnakoDB/artcat");
		ContentProviderClient artCatContProvCl = ctx.getContentResolver().acquireContentProviderClient(artCatAllURI);
		try
		{
			Cursor artCatCursor = artCatContProvCl.query(artCatAllURI, null, null, null, null);
			output[1] = this.getArtCatTableXLS(artCatCursor);
		} catch (RemoteException e)
		{
			e.printStackTrace();
		}

		//get ArtCatTable table and write it to xls
		Uri authorAllURI = Uri.parse("content://ru.kuchanov.odnako.db.ContentProviderOdnakoDB/author");
		ContentProviderClient authorContProvCl = ctx.getContentResolver().acquireContentProviderClient(authorAllURI);
		try
		{
			Cursor authorCursor = authorContProvCl.query(authorAllURI, null, null, null, null);
			output[2] = this.getAuthorXLS(authorCursor);
		} catch (RemoteException e)
		{
			e.printStackTrace();
		}

		//		output[0] = this.getArticleXLS();
		//		output[1] = this.getArtCatTableXLS();
		//		output[1] = this.getCategoryCSV();

		//		output[3] = this.getAuthorCSV();
		//		output[4] = this.getArtAutTableCSV();

		return output;
	}

	@Override
	protected void onPostExecute(File[] out)
	{
		Log.e(LOG, "onPostExecute");
		Intent emailIntent = new Intent(Intent.ACTION_SEND_MULTIPLE);
		emailIntent.setType("text/html");
		emailIntent.putExtra(android.content.Intent.EXTRA_EMAIL, new String[] { to });
		emailIntent.putExtra(android.content.Intent.EXTRA_SUBJECT, subj);
		emailIntent.putExtra(android.content.Intent.EXTRA_TEXT, msg);

		ArrayList<Uri> uris = new ArrayList<Uri>();
		for (File f : out)
		{
			uris.add(Uri.fromFile(f));
		}
		emailIntent.putParcelableArrayListExtra(Intent.EXTRA_STREAM, uris);

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
	private File getArtCatTableXLS(Cursor c)
	{
		try
		{
			File exportDir = new File(Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS),
			"");

			String path = exportDir.getCanonicalPath();

			//Instantiate a Workbook object.
			Workbook workbook = new Workbook();
			//Get the first worksheet's cells in the book.
			Cells cells = workbook.getWorksheets().get(0).getCells();

			//get column names from DB table
			String[] columnNames = getArtCatTableFieldsNames();
			//insert column names
			for (int i = 0; i < columnNames.length; i++)
			{
				cells.get(0, i).setValue(columnNames[i]);
			}

			//insert table data
			if (c.moveToFirst())
			{
				while (!c.isAfterLast())
				{
					for (int i = 0; i < c.getColumnCount(); i++)
					{
						String data = c.getString(c.getColumnIndex(columnNames[i]));
						cells.get(c.getPosition() + 1, i).setValue(data);
					}
					c.moveToNext();
				}
			}

			//Save the Excel file.
			workbook.save(path + "/ArtCatXLS.xls");

			File file = new File(path + "/ArtCatXLS.xls");

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

	//
	//	private File getCategoryCSV()
	//	{
	//		
	//	}

	private File getArticleXLS(Cursor c)
	{
		try
		{
			File exportDir = new File(Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS),
			"");

			String path = exportDir.getCanonicalPath();

			//Instantiate a Workbook object.
			Workbook workbook = new Workbook();
			//Get the first worksheet's cells in the book.
			Cells cells = workbook.getWorksheets().get(0).getCells();

			//get column names from DB table
			String[] columnNames = getArticleFieldsNames();
			//insert column names
			for (int i = 0; i < columnNames.length; i++)
			{
				cells.get(0, i).setValue(columnNames[i]);
			}

			//insert table data
			if (c.moveToFirst())
			{
				while (!c.isAfterLast())
				{
					for (int i = 0; i < c.getColumnCount(); i++)
					{
						String data = c.getString(c.getColumnIndex(columnNames[i]));
						cells.get(c.getPosition() + 1, i).setValue(data);
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

	private File getAuthorXLS(Cursor c)
	{
		try
		{
			File exportDir = new File(Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS),
			"");

			String path = exportDir.getCanonicalPath();

			//Instantiate a Workbook object.
			Workbook workbook = new Workbook();
			//Get the first worksheet's cells in the book.
			Cells cells = workbook.getWorksheets().get(0).getCells();

			//get column names from DB table
			String[] columnNames = getAuthorFieldsNames();
			//insert column names
			for (int i = 0; i < columnNames.length; i++)
			{
				cells.get(0, i).setValue(columnNames[i]);
			}

			//insert table data
			if (c.moveToFirst())
			{
				while (!c.isAfterLast())
				{
					for (int i = 0; i < c.getColumnCount(); i++)
					{
						String data = c.getString(c.getColumnIndex(columnNames[i]));
						cells.get(c.getPosition() + 1, i).setValue(data);
					}
					c.moveToNext();
				}
			}

			//Save the Excel file.
			workbook.save(path + "/AuthorXLS.xls");

			File file = new File(path + "/AuthorXLS.xls");

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
	 * returns String[] with names of all Article table columns
	 */
	public static String[] getArticleFieldsNames()
	{
		String[] arrStr1 = { "id", "url", "title", "img_art", "authorBlogUrl",
				"authorName", "preview", "pubDate", "refreshed", "numOfComments",
				"numOfSharings", "artText", "authorDescr", "tegs_main", "tegs_all",
				"share_quont", "to_read_main", "to_read_more", "img_author", "author" };
		return arrStr1;
	}

	/**
	 * returns String[] with names of all ArtCatTable table columns
	 */
	public static String[] getArtCatTableFieldsNames()
	{
		String[] arrStr1 = { "id", "article_id", "category_id", "nextArtUrl", "previousArtUrl", "isTop" };
		return arrStr1;
	}

	/**
	 * returns String[] with names of all Author table columns
	 */
	public static String[] getAuthorFieldsNames()
	{
		String[] arrStr1 = { "id", "blog_url", "name", "description", "who", "avatar", "avatarBig", "refreshed",
				"lastArticleDate", "firstArticleURL" };
		return arrStr1;
	}
}
