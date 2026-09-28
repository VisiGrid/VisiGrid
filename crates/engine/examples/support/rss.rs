//! Resident memory of this process, for the memory examples (#18).
//!
//! Shared by `visigrid-engine`'s `memsize` and `visigrid-io`'s `import_mem`
//! via `#[path]`. Linux reads `/proc/self/status`; macOS asks the kernel via
//! `task_info(MACH_TASK_BASIC_INFO)`. Elsewhere both return 0.

#![allow(dead_code)]

/// Current resident set size, in MB.
pub fn rss_mb() -> f64 {
    imp::rss_bytes() as f64 / (1024.0 * 1024.0)
}

/// Highest resident set size so far, in MB.
pub fn peak_rss_mb() -> f64 {
    imp::peak_rss_bytes() as f64 / (1024.0 * 1024.0)
}

/// Hand freed memory back to the OS where the allocator supports it, so RSS
/// after the call reflects what is live. glibc only; a no-op elsewhere.
pub fn trim_allocator() {
    #[cfg(all(target_os = "linux", target_env = "gnu"))]
    {
        extern "C" {
            fn malloc_trim(pad: usize) -> i32;
        }
        unsafe { malloc_trim(0) };
    }
}

#[cfg(target_os = "linux")]
mod imp {
    fn status_kb(field: &str) -> u64 {
        std::fs::read_to_string("/proc/self/status")
            .unwrap_or_default()
            .lines()
            .find_map(|line| line.strip_prefix(field))
            .and_then(|v| v.trim().trim_end_matches(" kB").parse::<u64>().ok())
            .unwrap_or(0)
    }

    pub fn rss_bytes() -> u64 {
        status_kb("VmRSS:") * 1024
    }

    pub fn peak_rss_bytes() -> u64 {
        status_kb("VmHWM:") * 1024
    }
}

#[cfg(target_os = "macos")]
mod imp {
    // <mach/task_info.h>: struct mach_task_basic_info, flavor 20.
    #[repr(C)]
    #[derive(Default)]
    struct MachTaskBasicInfo {
        virtual_size: u64,
        resident_size: u64,
        resident_size_max: u64,
        user_time: [i32; 2],
        system_time: [i32; 2],
        policy: i32,
        suspend_count: i32,
    }

    const MACH_TASK_BASIC_INFO: u32 = 20;
    const KERN_SUCCESS: i32 = 0;

    extern "C" {
        // `mach_task_self()` is a macro over this global.
        static mach_task_self_: u32;
        fn task_info(task: u32, flavor: u32, info: *mut i32, count: *mut u32) -> i32;
    }

    fn basic_info() -> Option<MachTaskBasicInfo> {
        let mut info = MachTaskBasicInfo::default();
        // Count is in natural_t (u32) units.
        let mut count = (std::mem::size_of::<MachTaskBasicInfo>() / 4) as u32;
        let kr = unsafe {
            task_info(
                mach_task_self_,
                MACH_TASK_BASIC_INFO,
                &mut info as *mut MachTaskBasicInfo as *mut i32,
                &mut count,
            )
        };
        (kr == KERN_SUCCESS).then_some(info)
    }

    pub fn rss_bytes() -> u64 {
        basic_info().map_or(0, |i| i.resident_size)
    }

    pub fn peak_rss_bytes() -> u64 {
        basic_info().map_or(0, |i| i.resident_size_max)
    }
}

#[cfg(not(any(target_os = "linux", target_os = "macos")))]
mod imp {
    pub fn rss_bytes() -> u64 {
        0
    }

    pub fn peak_rss_bytes() -> u64 {
        0
    }
}
