package ru.kuchanov.databaseexporter;


import ru.kuchanov.databaseexporter.db.XlsWriterTask;
import android.app.Activity;
import android.app.Fragment;
import android.content.Context;
import android.os.Bundle;
import android.util.Log;
import android.view.LayoutInflater;
import android.view.Menu;
import android.view.MenuItem;
import android.view.View;
import android.view.View.OnClickListener;
import android.view.ViewGroup;
import android.widget.Button;

public class ActivityMain extends Activity
{
	final static String TAG=ActivityMain.class.getSimpleName();

	@Override
	protected void onCreate(Bundle savedInstanceState)
	{
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_activity_main);
		if (savedInstanceState == null)
		{
			getFragmentManager().beginTransaction()
			.add(R.id.container, new PlaceholderFragment())
			.commit();
		}
		
//		Calendar calendar = Calendar.getInstance();
//	    calendar.set(Calendar.HOUR_OF_DAY, 14);
//	    calendar.set(Calendar.MINUTE, 0);
//	    calendar.set(Calendar.SECOND, 0);
//	    
////	    Log.e(TAG, calendar.toString());
//	    AlarmManager am = (AlarmManager) this.getSystemService(Context.ALARM_SERVICE);
//	    Intent intent = new Intent(this, ActivityMain.class);
//	    intent.setAction("notifyEveryDay");
//	    PendingIntent pendingIntent = PendingIntent.getBroadcast(this, 0, intent, PendingIntent.FLAG_UPDATE_CURRENT);
//	    am.setRepeating(AlarmManager.RTC, calendar.getTimeInMillis() ,AlarmManager.INTERVAL_DAY, pendingIntent);
	}

	@Override
	public boolean onCreateOptionsMenu(Menu menu)
	{
		// Inflate the menu; this adds items to the action bar if it is present.
		getMenuInflater().inflate(R.menu.activity_main, menu);
		return true;
	}

	@Override
	public boolean onOptionsItemSelected(MenuItem item)
	{
		// Handle action bar item clicks here. The action bar will
		// automatically handle clicks on the Home/Up button, so long
		// as you specify a parent activity in AndroidManifest.xml.
		int id = item.getItemId();
		if (id == R.id.action_settings)
		{
			return true;
		}
		return super.onOptionsItemSelected(item);
	}

	/**
	 * A placeholder fragment containing a simple view.
	 */
	public static class PlaceholderFragment extends Fragment
	{
		//tag:^(?!dalvikvm) tag:^(?!libEGL) tag:^(?!Open) tag:^(?!Google) tag:^(?!resour) tag:^(?!Chore) tag:^(?!EGL)
		private final static String TAG = PlaceholderFragment.class.getSimpleName();

		Context ctx;

		public PlaceholderFragment()
		{
		}

		@Override
		public void onCreate(Bundle b)
		{
			super.onCreate(b);

			this.ctx = this.getActivity();
		}

		@Override
		public View onCreateView(LayoutInflater inflater, ViewGroup container, Bundle savedInstanceState)
		{
			View v = inflater.inflate(R.layout.fragment_activity_main, container, false);

			Button b = (Button) v.findViewById(R.id.btn_1);
			b.setOnClickListener(new OnClickListener()
			{

				@Override
				public void onClick(View v)
				{
					Log.d(TAG, "btn_1 cliced");

					XlsWriterTask xlsWriter = new XlsWriterTask(ctx);
					xlsWriter.execute();
				}
			});

			return v;
		}
	}
}
