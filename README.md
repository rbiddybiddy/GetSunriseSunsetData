# GetSunriseSunsetData

It all started from someone wondering why the time of sunset seems to linger in one place for many days surrounding the
solstices, and then change rapidly from day to day in the spring and fall. Sounded like a sinusoid to me. Which
seemed plausible based on what I remembered from math & physics that circular and elliptical motion (like the earth's
orbit) tends to produce sinusoids in some dimension or other.

![simple harmonic motion](https://d-arora.github.io/VisualPhysics/mod5new/mod52B_files/image083.gif)

Now before you say "Yeah OK Copernicus!" Yes I know that this _groundbreaking theory_ of mine would be a very easy fact
to look up and confirm or refute. But what fun would that be, huh?

Instead, I went and found this free API that returns sunset & sunrise data for any date and location:\
[https://sunrise-sunset.org/api](https://sunrise-sunset.org/api)\
(It's a free service, so if you decide to use it, keep it to a reasonable volume.)

...and wrote this simple .NET console app GetSunriseSunsetData, which will:
1) Iterate through an entire year (or other interval), day by day
2) Query the above endpoint for sunrise & sunset times for each day
3) Dump the data into Excel where my skill level is such that I stand a decent chance of producing graphs that make sense.

So, do sunset and sunrise times follow a sine-like pattern? Spoiler alert: Yep. At least at my location in Oregon near
the 45ºN parallel. Although the curve seems skewed a bit for what I'm sure are _reasons_.

For writing Excel files, I used the NuGet package [ClosedXML](https://github.com/ClosedXML/ClosedXML), which is based on,
a wrapper for, slightly more straightforward than, and kind of a joke about, the
[Open XML](https://en.wikipedia.org/wiki/Office_Open_XML) spec that describes Excel files.

Configuration (e.g. the longitude and latitude for your location, what date to start on and how many days of data to
get) is done via private field initializations in the Program class. Of course these should be in a config file
or the like. Also I kind of deserialize the JSON the hard way. Look, this is a very simple and stupid app,
all right? There are better, easier and/or more sophisticated ways to do several of the things. Feel free to submit
improvements!