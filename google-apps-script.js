/**
 * Google Apps Script to populate Management Philosophy slides
 * 
 * Instructions:
 * 1. Go to script.google.com
 * 2. Create a new project
 * 3. Replace the default code with this script
 * 4. Update PRESENTATION_ID with your Google Slides ID
 * 5. Run the populateSlides function
 */

// Extract the presentation ID from your URL: 1xmY2iPUIQm0dbrGDWl-rLQDEr_kMDuH8rAvfL08mQyI
const PRESENTATION_ID = '1xmY2iPUIQm0dbrGDWl-rLQDEr_kMDuH8rAvfL08mQyI';

function populateSlides() {
  try {
    const presentation = SlidesApp.openById(PRESENTATION_ID);
    
    // Clear existing slides except the first one
    const slides = presentation.getSlides();
    for (let i = slides.length - 1; i > 0; i--) {
      slides[i].remove();
    }
    
    // Define slide content
    const slideData = [
      {
        title: "My Management Philosophy",
        subtitle: "Building teams that deliver customer impact through sustainable excellence"
      },
      {
        title: "What to Expect",
        bullets: [
          "My core values and how they guide decisions",
          "How I work and communicate",
          "What you can expect from me",
          "What I expect from you",
          "My approach to growth and development"
        ]
      },
      {
        title: "Four Principles That Guide Everything",
        bullets: [
          "Grow people through impactful work",
          "Build sustainable excellence",
          "Empower shared ownership",
          "Start with the customer, iterate with purpose"
        ]
      },
      {
        title: "Grow people through impactful work",
        bullets: [
          "Connect your career growth to meaningful customer & business outcomes",
          "Success = advancing your skills while delivering work that matters",
          "I'll help identify growth opportunities with clear strategic value"
        ],
        personalNote: "I get energized when the team is collaborating, driving their own fulfilling work, and connecting it to company strategy."
      },
      {
        title: "Build sustainable excellence",
        bullets: [
          "High performance without burnout",
          "Quality without perfectionism",
          "Work-life harmony while continuously improving",
          "Long hours are the exception, not the rule"
        ],
        personalNote: "We can bring our best selves to work by taking time off and signing off at normal hours."
      },
      {
        title: "Empower shared ownership",
        bullets: [
          "Total Football approach - anyone can contribute across responsibilities",
          "Distributed leadership with clear accountability", 
          "End-to-end ownership of your projects",
          "No silos, collaborative problem-solving"
        ],
        personalNote: "Teams are most effective when we distribute and share responsibilities while empowering ownership.",
        imageUrl: "https://static.independent.co.uk/2021/07/22/12/newFile-12.jpg?width=1200"
      },
      {
        title: "Start with the customer, iterate with purpose",
        bullets: [
          "Begin with customer needs, work backwards to solutions",
          "Build only what's needed - avoid premature optimization",
          "Transparent SMART goals everyone can see and contribute to",
          "Continuous feedback loops for improvement"
        ],
        personalNote: "Connect our work to customer significance and business impact so we feel purpose."
      },
      {
        title: "How I Communicate",
        subtitle: "Direct, Transparent, Collaborative",
        bullets: [
          "Continuous feedback - both positive and constructive",
          "No surprises - you'll always know how you're performing",
          "Open to input - I actively seek feedback on my decisions",
          "Strategic context - understand the 'why' behind your work"
        ]
      },
      {
        title: "What You Can Expect From Me",
        subtitle: "My Commitments to You",
        bullets: [
          "Career development support and growth planning",
          "Blocker removal and stakeholder management",
          "Strategic context connecting work to customer impact",
          "Advocacy for technical investment alongside features",
          "Protection from unrealistic timelines and unsustainable scope"
        ]
      },
      {
        title: "What I Expect From You",
        subtitle: "Partnership for Success",
        bullets: [
          "End-to-end ownership of your projects",
          "Quality focus - core experience and high-impact scenarios handled well",
          "Continuous learning and professional growth",
          "Transparent communication about blockers, risks, decisions",
          "Customer awareness of how your work creates impact"
        ]
      },
      {
        title: "Technical Philosophy",
        subtitle: "Simple, Customer-Focused Excellence",
        bullets: [
          "Incremental improvement - invest in technical debt alongside features",
          "Simple, essential solutions without unnecessary complexity",
          "Architecture that enables faster customer value delivery",
          "Quality without perfectionism - iterate based on real feedback"
        ]
      },
      {
        title: "Growth & Development",
        subtitle: "How We'll Help You Grow",
        bullets: [
          "Stretch assignments that push your skills with support",
          "Cross-functional exposure to broader business context",
          "Technical depth and breadth development",
          "Leadership opportunities to mentor and influence",
          "Regular career conversations about your goals"
        ]
      },
      {
        title: "Questions & Feedback",
        subtitle: "Always Open for Conversation",
        bullets: [
          "Ask for clarification on expectations or priorities",
          "Share feedback on how I can better support you",
          "Raise concerns about team dynamics or processes",
          "Suggest improvements to how we work together"
        ],
        personalNote: "Remember: My job is to help you succeed"
      },
      {
        title: "Let's Build Something Great Together",
        bullets: [
          "Authentic leadership through clear values",
          "Sustainable excellence with customer impact",
          "Shared ownership and continuous growth",
          "Open communication and collaboration"
        ],
        subtitle: "Questions?"
      }
    ];
    
    // Create slides
    slideData.forEach((slideContent, index) => {
      let slide;
      if (index === 0) {
        // Use the first existing slide
        slide = slides[0];
        slide.getShapes().forEach(shape => shape.remove());
      } else {
        // Create new slide
        slide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
      }
      
      createSlideContent(slide, slideContent);
    });
    
    console.log('Slides populated successfully!');
    console.log(`View your presentation: https://docs.google.com/presentation/d/${PRESENTATION_ID}/edit`);
    
  } catch (error) {
    console.error('Error populating slides:', error);
  }
}

function createSlideContent(slide, content) {
  // Clear existing shapes first to avoid overlap
  const shapes = slide.getShapes();
  shapes.forEach(shape => {
    try {
      shape.remove();
    } catch (e) {
      // Some shapes might not be removable, that's okay
    }
  });
  
  // Set a subtle background color
  try {
    slide.getBackground().setSolidFill('#f8f9fa');
  } catch (e) {
    // Background setting might not be supported in all contexts
  }
  
  // Standard slide dimensions (assuming 16:9 aspect ratio)
  const slideWidth = 720;
  const slideHeight = 405;
  const margin = 40;
  
  // Title area
  const titleTop = 30;
  const titleHeight = 80;
  
  // Content area starts after title
  const contentTop = titleTop + titleHeight + 20;
  const contentHeight = slideHeight - contentTop - margin;
  
  // Create title
  if (content.title) {
    const titleShape = slide.insertTextBox(content.title);
    const titleRange = titleShape.getText();
    titleRange.getTextStyle()
      .setFontSize(36)
      .setBold(true)
      .setForegroundColor('#2c3e50');
    
    // Position title at top of slide
    titleShape.setTop(titleTop);
    titleShape.setLeft(margin);
    titleShape.setWidth(slideWidth - (margin * 2));
    titleShape.setHeight(titleHeight);
  }
  
  // Build and create content
  if (content.subtitle || content.bullets || content.personalNote || content.imageUrl) {
    let bodyContent = '';
    
    if (content.subtitle) {
      bodyContent += content.subtitle + '\n\n';
    }
    
    if (content.bullets) {
      bodyContent += content.bullets.map(bullet => `• ${bullet}`).join('\n');
      if (content.personalNote) bodyContent += '\n\n';
    }
    
    if (content.personalNote) {
      bodyContent += content.personalNote;
    }
    
    // Adjust content area if image is present
    let textWidth = slideWidth - (margin * 2);
    let imageWidth = 0;
    
    if (content.imageUrl) {
      imageWidth = 200; // Image width
      textWidth = slideWidth - (margin * 3) - imageWidth; // Text area width
      
      try {
        // Add image to the right side
        const imageBlob = UrlFetchApp.fetch(content.imageUrl).getBlob();
        const image = slide.insertImage(imageBlob);
        
        // Position image on the right
        image.setTop(contentTop);
        image.setLeft(slideWidth - margin - imageWidth);
        image.setWidth(imageWidth);
        image.setHeight(150); // Fixed height
      } catch (e) {
        console.log('Could not load image:', e);
        // If image fails, use full width for text
        textWidth = slideWidth - (margin * 2);
      }
    }
    
    // Create main content box
    if (bodyContent) {
      const bodyShape = slide.insertTextBox(bodyContent);
      const bodyRange = bodyShape.getText();
      
      // Position content area (adjusted for image if present)
      bodyShape.setTop(contentTop);
      bodyShape.setLeft(margin);
      bodyShape.setWidth(textWidth);
      bodyShape.setHeight(contentHeight);
      
      // Style the body text (default)
      bodyRange.getTextStyle()
        .setFontSize(18)
        .setForegroundColor('#34495e');
      
      // Style subtitle if present
      if (content.subtitle) {
        const subtitleEnd = content.subtitle.length;
        bodyRange.getRange(0, subtitleEnd).getTextStyle()
          .setFontSize(24)
          .setItalic(true)
          .setForegroundColor('#7f8c8d');
      }
      
      // Style personal note if present
      if (content.personalNote) {
        const noteStart = bodyContent.lastIndexOf(content.personalNote);
        if (noteStart >= 0) {
          bodyRange.getRange(noteStart, noteStart + content.personalNote.length).getTextStyle()
            .setFontSize(16)
            .setItalic(true)
            .setForegroundColor('#3498db');
        }
      }
    }
  }
}

function createBodyContent(slide, content, startY) {
  let currentY = startY;
  
  // Add subtitle
  if (content.subtitle) {
    const subtitleShape = slide.insertTextBox(content.subtitle);
    const subtitleRange = subtitleShape.getText();
    subtitleRange.getTextStyle()
      .setFontSize(24)
      .setItalic(true)
      .setForegroundColor('#666666');
    
    subtitleShape.setTop(currentY);
    subtitleShape.setLeft(50);
    subtitleShape.setWidth(600);
    subtitleShape.setHeight(50);
    currentY += 70;
  }
  
  // Add bullets
  if (content.bullets) {
    const bulletText = content.bullets.map(bullet => `• ${bullet}`).join('\n');
    const bulletShape = slide.insertTextBox(bulletText);
    const bulletRange = bulletShape.getText();
    bulletRange.getTextStyle()
      .setFontSize(18)
      .setForegroundColor('#34495e');
    
    bulletShape.setTop(currentY);
    bulletShape.setLeft(50);
    bulletShape.setWidth(600);
    bulletShape.setHeight(Math.max(200, content.bullets.length * 30));
    currentY += Math.max(200, content.bullets.length * 30) + 20;
  }
  
  // Add quote
  if (content.personalNote) {
    const noteShape = slide.insertTextBox(content.personalNote);
    const noteRange = noteShape.getText();
    noteRange.getTextStyle()
      .setFontSize(16)
      .setItalic(true)
      .setForegroundColor('#0066cc');
    
    noteShape.setTop(currentY);
    noteShape.setLeft(50);
    noteShape.setWidth(600);
    noteShape.setHeight(60);
  }
}

function createSlideContentFallback(slide, content) {
  // Original method as fallback
  if (content.title) {
    const titleShape = slide.insertTextBox(content.title);
    const titleRange = titleShape.getText();
    titleRange.getTextStyle()
      .setFontSize(36)
      .setBold(true)
      .setForegroundColor('#2c3e50');
    
    titleShape.setTop(50);
    titleShape.setLeft(50);
    titleShape.setWidth(600);
    titleShape.setHeight(80);
  }
  
  let currentY = 150;
  
  if (content.subtitle) {
    const subtitleShape = slide.insertTextBox(content.subtitle);
    const subtitleRange = subtitleShape.getText();
    subtitleRange.getTextStyle()
      .setFontSize(24)
      .setItalic(true)
      .setForegroundColor('#666666');
    
    subtitleShape.setTop(currentY);
    subtitleShape.setLeft(50);
    subtitleShape.setWidth(600);
    subtitleShape.setHeight(50);
    currentY += 70;
  }
  
  if (content.bullets) {
    const bulletText = content.bullets.map(bullet => `• ${bullet}`).join('\n');
    const bulletShape = slide.insertTextBox(bulletText);
    const bulletRange = bulletShape.getText();
    bulletRange.getTextStyle()
      .setFontSize(18)
      .setForegroundColor('#34495e');
    
    bulletShape.setTop(currentY);
    bulletShape.setLeft(50);
    bulletShape.setWidth(600);
    bulletShape.setHeight(Math.max(200, content.bullets.length * 30));
    currentY += Math.max(200, content.bullets.length * 30) + 20;
  }
  
  if (content.personalNote) {
    const noteShape = slide.insertTextBox(content.personalNote);
    const noteRange = noteShape.getText();
    noteRange.getTextStyle()
      .setFontSize(16)
      .setItalic(true)
      .setForegroundColor('#0066cc');
    
    noteShape.setTop(currentY);
    noteShape.setLeft(50);
    noteShape.setWidth(600);
    noteShape.setHeight(60);
  }
}

// Helper function to test the script
function testConnection() {
  try {
    const presentation = SlidesApp.openById(PRESENTATION_ID);
    console.log('Successfully connected to presentation:', presentation.getName());
    console.log('Current number of slides:', presentation.getSlides().length);
  } catch (error) {
    console.error('Error connecting to presentation:', error);
    console.log('Make sure the PRESENTATION_ID is correct and you have edit access to the presentation.');
  }
}